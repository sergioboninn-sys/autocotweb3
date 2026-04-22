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

# --- ABA 1: COTAÇÃO (CORREÇÃO APLICADA) ---
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
            
            # Localizar índices das colunas
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            try:
                d_idx = col_indices[desc_col.strip()]
                b_idx = col_indices[bar_col.strip()]
                p_idx = col_indices[price_col.strip()]
                
                # ZERAR O CONTADOR PARA CADA PROCESSAMENTO
                total_preenchidos = 0
                
                # Percorrer as linhas
                for r in range(int(start_row), ws.max_row + 1):
                    desc_cell_val = ws.cell(row=r, column=d_idx).value
                    
                    # Se não houver descrição, ignoramos a linha (limpeza de linhas fantasmas)
                    if not desc_cell_val or str(desc_cell_val).strip() == "":
                        continue
                    
                    found_price = None
                    bar_cell_val = ws.cell(row=r, column=b_idx).value
                    
                    # 1. Busca por Barras
                    if "Barras" in modo or "Híbrido" in modo:
                        eans = extract_all_barcodes(bar_cell_val)
                        for ean in eans:
                            if ean in price_map:
                                found_price = price_map[ean]
                                break
                    
                    # 2. Busca por Similaridade
                    if found_price is None and ("Similaridade" in modo or "Híbrido" in modo):
                        best_sim = 0
                        d_det = extrair_detalhes(desc_cell_val)
                        for _, db_row in master_db.iterrows():
                            sim = similarity(desc_cell_val, db_row['Description'])
                            if sim >= 0.75 and d_det == extrair_detalhes(db_row['Description']):
                                if sim > best_sim:
                                    best_sim = sim
                                    found_price = db_row['Price']
                    
                    # 3. Escrita e Contagem (SÓ CONTA SE ENCONTRAR O PREÇO)
                    if found_price is not None:
                        calc_price = float(found_price) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(calc_price) if aplicar_arredondamento else calc_price
                        total_preenchidos += 1

                # Salvar em memória
                out = io.BytesIO()
                wb.save(out)
                
                st.success(f"Sucesso! Foram preenchidos **{total_preenchidos}** preços na planilha.")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "cotacao_corrigida.xlsx")
                
            except Exception as e:
                st.error(f"Erro ao processar colunas: {e}")

# --- ABA 2: VENDAS (ORIGINAL) ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("🔍 Busca e Adição Rápida")
        query = st.text_input("Pesquisar produto:", placeholder="Digite e pressione Enter")
        desc_geral = st.number_input("Desconto Padrão (%)", 0.0)
        
        if query:
            terms = query.upper().split()
            mask = master_db['Description'].apply(lambda x: all(t in str(x).upper() for t in terms))
            results = master_db[mask].head(20)
            
            for idx, row in results.iterrows():
                p_sug = extra_round(float(row['Price']) * (1 - (desc_geral/100)))
                with st.container():
                    c1, c2, c3, c4 = st.columns([3, 1, 1, 0.5])
                    c1.write(row['Description'])
                    c2.write(f"R$ {p_sug}")
                    qtd = c3.number_input("Qtd", 1, 1000, 1, key=f"q_{idx}")
                    if c4.button("➕", key=f"b_{idx}"):
                        st.session_state.carrinho.append({
                            "Descrição": row['Description'], "EAN": row['Barcode'],
                            "Qtd": qtd, "Preço": p_sug, "Total": extra_round(qtd * p_sug)
                        })
                        st.rerun()

    with col2:
        st.subheader("🛒 Carrinho")
        if st.session_state.carrinho:
            df_c = pd.DataFrame(st.session_state.carrinho)
            st.dataframe(df_c[["Descrição", "Qtd", "Total"]], hide_index=True)
            st.metric("Total", f"R$ {extra_round(df_c['Total'].sum())}")
            if st.button("Limpar"):
                st.session_state.carrinho = []
                st.rerun()
        else:
            st.info("Vazio")

# --- ABA 3: BANCO (ORIGINAL) ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco")
    f = st.file_uploader("Upload Banco (xlsx/csv)", type=["xlsx", "csv"])
    if f and st.button("Salvar Banco"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['Description', 'Barcode', 'Price']
        df['Barcode'] = df['Barcode'].astype(str).str.replace(r'\.0$', '', regex=True)
        df.to_csv(DB_STORAGE, index=False)
        st.success("Banco salvo com sucesso!")

# --- ABA 4: USUÁRIOS (ORIGINAL) ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    st.write("Usuários Cadastrados:")
    st.write(pd.DataFrame.from_dict(users, orient='index'))
