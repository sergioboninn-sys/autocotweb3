import streamlit as st
import pandas as pd
import io
import re
import json
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from streamlit.components.v1 import html

# --- CONFIGURAÇÃO E PERSISTÊNCIA ---
st.set_page_config(page_title="Automatizador de Preços PRO", layout="wide")

# --- INJEÇÃO DE JAVASCRIPT (F1 E F5) ---
js_shortcuts = """
<script>
    const doc = window.parent.document;
    doc.addEventListener('keydown', function(e) {
        if (e.key === 'F1') {
            e.preventDefault();
            const inputs = doc.querySelectorAll('input[type="text"]');
            const searchInput = Array.from(inputs).find(el => el.placeholder.includes("Digite e pressione"));
            if (searchInput) searchInput.focus();
        }
        if (e.key === 'F5') {
            e.preventDefault();
            const inputs = doc.querySelectorAll('input[type="text"]');
            const searchInput = Array.from(inputs).find(el => el.placeholder.includes("Digite e pressione"));
            if (searchInput) {
                searchInput.value = "";
                searchInput.dispatchEvent(new Event('input', { bubbles: true }));
                searchInput.dispatchEvent(new Event('change', { bubbles: true }));
                searchInput.focus();
            }
        }
    });
</script>
"""
html(js_shortcuts, height=0)

DB_STORAGE = "master_database.csv"
USERS_STORAGE = "users_db.json"

if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []

if 'replicar_data' not in st.session_state:
    st.session_state.replicar_data = None

def load_users():
    if not os.path.exists(USERS_STORAGE):
        return {"admin": {"password": "admin123", "expiry": "2099-12-31", "role": "admin"}}
    with open(USERS_STORAGE, "r") as f: return json.load(f)

def save_users(users):
    with open(USERS_STORAGE, "w") as f: json.dump(users, f)

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def login():
    st.sidebar.title("🔐 Acesso Restrito")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        users = load_users()
        if u in users and users[u]['password'] == p:
            exp = datetime.strptime(users[u]['expiry'], "%Y-%m-%d")
            if datetime.now() <= exp:
                st.session_state.autenticado = True
                st.session_state.user_role = users[u]['role']
                st.rerun()
        st.sidebar.error("Acesso negado")

if not st.session_state.autenticado:
    login()
    st.stop()

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if pd.isna(valor): return valor
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def extract_all_barcodes(val):
    if pd.isna(val) or val is None: return []
    text = str(val).split('.')[0]
    return re.findall(r'\d{8,14}', text)

def similarity(a, b):
    return SequenceMatcher(None, str(a).upper().strip(), str(b).upper().strip()).ratio()

def extrair_detalhes(texto):
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', str(texto).lower()))

# --- CARREGAMENTO DO BANCO ---
@st.cache_data
def get_master_db():
    if os.path.exists(DB_STORAGE):
        df = pd.read_csv(DB_STORAGE)
        df['Barcode'] = df['Barcode'].astype(str).str.replace(r'\.0$', '', regex=True)
        return df
    return pd.DataFrame(columns=['Description', 'Barcode', 'Price'])

# --- INTERFACE ---
tabs = ["📊 Cotação", "💰 Vendas", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO (CORREÇÃO DEFINITIVA) ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido (Barras + Similaridade)", "Apenas Barras", "Apenas Similaridade"])
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Planilha de Destino", type=["xlsx"])
    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início dos produtos:", 1, 1000, 11)
        
        t_df_view = pd.read_excel(target_file, header=header_pos-1)
        col1, col2, col3 = st.columns(3)
        desc_col = col1.selectbox("Coluna Descrição", t_df_view.columns)
        bar_col = col2.selectbox("Coluna Barras", t_df_view.columns)
        price_col = col3.selectbox("Coluna Preço", t_df_view.columns)

        if st.button("🚀 Processar e Preservar Formatação"):
            price_map = dict(zip(master_db['Barcode'].astype(str), master_db['Price']))
            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            # Mapeamento de colunas
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            try:
                d_idx, b_idx, p_idx = col_indices[desc_col.strip()], col_indices[bar_col.strip()], col_indices[price_col.strip()]
                
                # 1. MEMORIZAR ESTADO INICIAL (Apenas linhas que têm descrição)
                vazios_antes = []
                for r in range(int(start_row), ws.max_row + 1):
                    # Se a linha não tem descrição, ignoramos completamente (evita contar o lixo no fim da planilha)
                    if not ws.cell(row=r, column=d_idx).value or str(ws.cell(row=r, column=d_idx).value).strip() == "":
                        continue
                    
                    val_p = ws.cell(row=r, column=p_idx).value
                    if val_p is None or str(val_p).strip() == "":
                        vazios_antes.append(r)

                # 2. PROCESSAR
                for r in vazios_antes:
                    d_val = ws.cell(row=r, column=d_idx).value
                    found_p = None
                    
                    # Busca por Barras
                    if "Barras" in modo or "Híbrido" in modo:
                        for b in extract_all_barcodes(ws.cell(row=r, column=b_idx).value):
                            if b in price_map:
                                found_p = price_map[b]
                                break
                    
                    # Busca por Similaridade
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo):
                        best_sim = 0
                        d_det = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim = similarity(d_val, row_db['Description'])
                            if sim >= 0.75 and d_det == extrair_detalhes(row_db['Description']):
                                if sim > best_sim:
                                    best_sim = sim
                                    found_p = row_db['Price']
                    
                    if found_p is not None:
                        f_p = float(found_p) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(f_p) if aplicar_arredondamento else f_p

                # 3. VERIFICAR A DIFERENÇA (QUANTOS ERAM VAZIOS E AGORA TÊM PREÇO)
                total_preenchidos = 0
                for r in vazios_antes:
                    val_atual = ws.cell(row=r, column=p_idx).value
                    if val_atual is not None and str(val_atual).strip() != "":
                        total_preenchidos += 1

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Sucesso! Foram preenchidos **{total_preenchidos}** itens que estavam sem preço.")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "cotacao_final.xlsx")
                
            except Exception as e:
                st.error(f"Erro no processamento: {e}")

# --- ABA 2: VENDAS (LAYOUT E FUNÇÕES ORIGINAIS RESTAURADAS) ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    col_vendas_1, col_vendas_2 = st.columns([3, 1])
    
    with col_vendas_1:
        st.subheader("🔍 Busca e Adição Rápida")
        query = st.text_input("Pesquisar produto:", placeholder="Digite e pressione Enter", key="search_query")
        desc_geral = st.number_input("Desconto Padrão na Tabela (%)", 0.0, 100.0, 0.0)
        
        if query:
            search_terms = query.replace('%', ' ').split()
            mask = master_db['Description'].apply(lambda x: all(term.upper() in str(x).upper() for term in search_terms))
            results = master_db[mask].head(20).copy() 
            
            if not results.empty:
                if st.session_state.replicar_data:
                    with st.container():
                        st.warning("🔄 **Replicação de Família Detectada**")
                        rep = st.session_state.replicar_data
                        st.write(f"Deseja replicar Qtd: {rep['qtd']} e Preço: R$ {rep['preco']} para itens da família '{rep['familia']}'?")
                        familia_results = results[results['Description'].str.contains(rep['familia'], case=False) & (results['Barcode'] != rep['ean'])]
                        
                        if st.button("Confirmar em Todos da Família"):
                            for _, rf in familia_results.iterrows():
                                st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                            st.session_state.replicar_data = None
                            st.rerun()
                
                for idx, row in results.iterrows():
                    p_sugestao = extra_round(float(row['Price']) * (1 - (desc_geral/100)))
                    with st.container():
                        c_desc, c_ean, c_preco, c_qtd, c_add = st.columns([3, 1.5, 1.2, 1, 0.5])
                        c_desc.write(f"**{row['Description']}**")
                        c_ean.write(f"`{row['Barcode']}`")
                        c_preco.write(f"R$ {p_sugestao}")
                        input_qtd = c_qtd.number_input("Qtd", 1, 1000, 1, key=f"qtd_{idx}")
                        if c_add.button("➕", key=f"btn_{idx}"):
                            st.session_state.carrinho.append({"Descrição": row['Description'], "EAN": row['Barcode'], "Qtd": input_qtd, "Preço Unit": p_sugestao, "Total": extra_round(p_sugestao * input_qtd)})
                            palavras = row['Description'].split()
                            if len(palavras) >= 2:
                                st.session_state.replicar_data = {"familia": f"{palavras[0]} {palavras[1]}", "qtd": input_qtd, "preco": p_sugestao, "ean": row['Barcode']}
                            st.rerun()

    with col_vendas_2:
        st.subheader("🛒 Seu Pedido")
        if st.session_state.carrinho:
            df_cart = pd.DataFrame(st.session_state.carrinho)
            for item in st.session_state.carrinho:
                st.write(f"{item['Qtd']}x {item['Descrição']} - **R$ {item['Total']}**")
            st.metric("Total Geral", f"R$ {extra_round(df_cart['Total'].sum())}")
            if st.button("🗑️ Limpar Tudo"):
                st.session_state.carrinho = []
                st.rerun()
            output_xlsx = io.BytesIO()
            df_cart.to_excel(output_xlsx, index=False)
            st.download_button("📥 Baixar Pedido", output_xlsx.getvalue(), "pedido.xlsx")

# --- AS OUTRAS ABAS (BANCO E USUÁRIOS) SEGUEM O PADRÃO ORIGINAL ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco")
    f = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    if f and st.button("💾 Salvar Banco"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df.columns = ['Description', 'Barcode', 'Price']
        df.to_csv(DB_STORAGE, index=False)
        st.success("Banco Atualizado!")

elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    st.write(pd.DataFrame.from_dict(users, orient='index'))
