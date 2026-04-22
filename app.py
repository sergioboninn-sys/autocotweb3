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

# --- ABA 1: COTAÇÃO (CORREÇÃO FILTRADA) ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Planilha de Destino", type=["xlsx"])
    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início:", 1, 1000, 11)
        
        t_df_view = pd.read_excel(target_file, header=header_pos-1)
        col1, col2, col3 = st.columns(3)
        desc_col = col1.selectbox("Coluna Descrição", t_df_view.columns)
        bar_col = col2.selectbox("Coluna Barras", t_df_view.columns)
        price_col = col3.selectbox("Coluna Preço", t_df_view.columns)

        if st.button("🚀 Processar"):
            price_map = dict(zip(master_db['Barcode'].astype(str), master_db['Price']))
            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            # Mapear colunas
            col_map = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            d_idx, b_idx, p_idx = col_map[desc_col], col_map[bar_col], col_map[price_col]
            
            # CONTAGEM REAL
            contador_sucesso = 0

            for r in range(int(start_row), ws.max_row + 1):
                desc_val = ws.cell(row=r, column=d_idx).value
                
                # SE A DESCRIÇÃO ESTIVER TOTALMENTE VAZIA, PULA (IGNORA LINHAS FANTASMAS)
                if desc_val is None or str(desc_val).strip() == "":
                    continue

                found_price = None
                bar_val = ws.cell(row=r, column=b_idx).value
                
                # Busca
                if "Barras" in modo or "Híbrido" in modo:
                    eans = extract_all_barcodes(bar_val)
                    for e in eans:
                        if e in price_map:
                            found_price = price_map[e]
                            break
                
                if found_price is None and ("Similaridade" in modo or "Híbrido" in modo):
                    best_sim = 0
                    d_det = extrair_detalhes(desc_val)
                    for _, db_row in master_db.iterrows():
                        sim = similarity(desc_val, db_row['Description'])
                        if sim >= 0.75 and d_det == extrair_detalhes(db_row['Description']):
                            if sim > best_sim:
                                best_sim = sim
                                found_price = db_row['Price']

                # SÓ GRAVA E CONTA SE REALMENTE ENCONTROU O PREÇO
                if found_price is not None:
                    final_p = float(found_price) * (1 - (discount/100))
                    ws.cell(row=r, column=p_idx).value = extra_round(final_p) if aplicar_arredondamento else final_p
                    contador_sucesso += 1

            out = io.BytesIO()
            wb.save(out)
            st.success(f"Feito! **{contador_sucesso}** itens foram atualizados.")
            st.download_button("Baixar Resultado", out.getvalue(), "resultado.xlsx")

# --- AS OUTRAS ABAS (VENDAS, BANCO, USUÁRIOS) FORAM MANTIDAS INTEGRALMENTE ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    col1, col2 = st.columns([3, 1])
    with col1:
        query = st.text_input("Pesquisar produto:", key="vendas_search")
        if query:
            terms = query.upper().split()
            mask = master_db['Description'].apply(lambda x: all(t in str(x).upper() for t in terms))
            results = master_db[mask].head(20)
            for idx, row in results.iterrows():
                with st.container():
                    c1, c2, c3 = st.columns([3, 1, 1])
                    c1.write(row['Description'])
                    c2.write(f"R$ {row['Price']}")
                    if c3.button("➕", key=f"add_{idx}"):
                        st.session_state.carrinho.append({"item": row['Description'], "preco": row['Price']})
                        st.rerun()
    with col2:
        st.subheader("🛒 Carrinho")
        for i in st.session_state.carrinho:
            st.write(f"- {i['item']}")

elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco")
    f = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    if f and st.button("Salvar Banco"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df.to_csv(DB_STORAGE, index=False)
        st.success("Banco Atualizado!")

elif aba == "👤 Usuários":
    st.title("👤 Usuários")
    users = load_users()
    st.write(users)
