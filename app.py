import streamlit as st
import pandas as pd
import io
import re
import json
import os
import openpyxl
from openpyxl.utils import get_column_letter
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

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    if master_db.empty: st.warning("Banco vazio.")
    
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
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            try:
                d_idx, b_idx, p_idx = col_indices[desc_col.strip()], col_indices[bar_col.strip()], col_indices[price_col.strip()]
                linhas_com_preco = []

                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = ws.cell(row=r, column=b_idx).value
                    
                    if not d_val or str(d_val).strip() == "":
                        continue
                        
                    found_p = None
                    if "Barras" in modo or "Híbrido" in modo:
                        for b in extract_all_barcodes(b_val):
                            if b in price_map: 
                                found_p = price_map[b]
                                break
                    
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
                        linhas_com_preco.append(r)

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Concluído! Total de itens com preços preenchidos: **{len(linhas_com_preco)}**")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "cotacao_final.xlsx")
            except Exception as e: 
                st.error(f"Erro: {e}")

# --- ABA 2: VENDAS ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    col_vendas_1, col_vendas_2 = st.columns([3, 1])
    
    with col_vendas_1:
        st.subheader("🔍 Busca e Adição Rápida")
        query = st.text_input("Pesquisar produto:", placeholder="Digite e pressione Enter", key="search_query")
        desc_geral = st.number_input("Desconto (%)", 0.0, 100.0, 0.0)
        
        if query:
            search_terms = query.replace('%', ' ').split()
            mask = master_db['Description'].apply(lambda x: all(term.upper() in str(x).upper() for term in search_terms))
            results = master_db[mask].head(20).copy() 
            
            if not results.empty:
                if st.session_state.replicar_data:
                    with st.container():
                        st.warning("🔄 **Replicação de Família**")
                        rep = st.session_state.replicar_data
                        st.write(f"Replicar Qtd: {rep['qtd']} e Preço: R$ {rep['preco']} para '{rep['familia']}'?")
                        familia_results = results[results['Description'].str.contains(rep['familia'], case=False) & (results['Barcode'] != rep['ean'])]
                        
                        if st.button("Confirmar em Todos"):
                            for _, rf in familia_results.iterrows():
                                st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                            st.session_state.replicar_data = None
                            st.rerun()
                
                for idx, row in results.iterrows():
                    p_sug = extra_round(float(row['Price']) * (1 - (desc_geral/100)))
                    with st.container():
                        c_desc, c_ean, c_preco, c_qtd, c_add = st.columns([3, 1.5, 1.2, 1, 0.5])
                        c_desc.write(f"**{row['Description']}**")
                        c_ean.write(f"`{row['Barcode']}`")
                        c_preco.write(f"R$ {p_sug}")
                        iqtd = c_qtd.number_input("Qtd", 1, 1000, 1, key=f"v_{idx}")
                        if c_add.button("➕", key=f"b_{idx}"):
                            st.session_state.carrinho.append({"Descrição": row['Description'], "EAN": row['Barcode'], "Qtd": iqtd, "Preço Unit": p_sug, "Total": extra_round(p_sug * iqtd)})
                            pals = row['Description'].split()
                            if len(pals) >= 2:
                                st.session_state.replicar_data = {"familia": f"{pals[0]} {pals[1]}", "qtd": iqtd, "preco": p_sug, "ean": row['Barcode']}
                            st.rerun()

    with col_vendas_2:
        st.subheader("🛒 Pedido")
        if st.session_state.carrinho:
            df_c = pd.DataFrame(st.session_state.carrinho)
            for i in st.session_state.carrinho:
                st.write(f"{i['Qtd']}x {i['Descrição']} - R$ {i['Total']}")
            st.metric("Total", f"R$ {extra_round(df_c['Total'].sum())}")
            if st.button("🗑️ Limpar"):
                st.session_state.carrinho = []
                st.rerun()
            output = io.BytesIO()
            df_c.to_excel(output, index=False)
            st.download_button("📥 Excel", output.getvalue(), "pedido.xlsx")

# --- OUTRAS ABAS ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Banco de Dados")
    f = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    if f and st.button("💾 Salvar"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['Description', 'Barcode', 'Price']
        df['Barcode'] = df['Barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
        df.to_csv(DB_STORAGE, index=False)
        st.cache_data.clear()
        st.success("Atualizado!")

elif aba == "👤 Usuários":
    st.title("👤 Usuários")
    users = load_users()
    st.write(pd.DataFrame.from_dict(users, orient='index'))
