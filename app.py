import streamlit as st
import pandas as pd
import io
import re
import json
import os
import openpyxl
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher

# --- CONFIGURAÇÃO E PERSISTÊNCIA ---
st.set_page_config(page_title="Automatizador de Preços PRO", layout="wide")

DB_STORAGE = "master_database.csv"
USERS_STORAGE = "users_db.json"

if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []

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

# --- ABA 1: COTAÇÃO (MANTIDA) ---
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
                count = 0
                for r in range(int(start_row), ws.max_row + 1):
                    d_val, b_val = ws.cell(row=r, column=d_idx).value, ws.cell(row=r, column=b_idx).value
                    found_p = None
                    if "Barras" in modo or "Híbrido" in modo:
                        for b in extract_all_barcodes(b_val):
                            if b in price_map: found_p = price_map[b]; break
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo) and d_val:
                        best_sim = 0
                        d_det = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim = similarity(d_val, row_db['Description'])
                            if sim >= 0.75 and d_det == extrair_detalhes(row_db['Description']):
                                if sim > best_sim: best_sim = sim; found_p = row_db['Price']
                    if found_p:
                        f_p = float(found_p) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(f_p) if aplicar_arredondamento else f_p
                        count += 1
                out = io.BytesIO()
                wb.save(out)
                st.success(f"{count} itens preenchidos!")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "cotacao_final.xlsx")
            except Exception as e: st.error(f"Erro: {e}")

# --- ABA 2: SISTEMA DE VENDAS (NOVA) ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    
    col_c1, col_c2 = st.columns([2, 1])
    
    with col_c1:
        st.subheader("🔍 Busca de Produtos")
        query = st.text_input("Produto (Ex: 'ref tang' ou 'deter%ipe')", placeholder="Digite parte do nome...")
        desc_geral = st.number_input("Desconto Geral na Tabela (%)", 0.0, 100.0, 0.0)
        
        if query:
            # Lógica de Busca Avançada
            search_terms = query.replace('%', ' ').split()
            mask = master_db['Description'].apply(lambda x: all(term.upper() in str(x).upper() for term in search_terms))
            results = master_db[mask].copy()
            
            if not results.empty:
                results['Preço Tabela'] = results['Price']
                results['Preço c/ Desc'] = results['Price'] * (1 - (desc_geral/100))
                results['Preço c/ Desc'] = results['Preço c/ Desc'].apply(extra_round)
                
                st.dataframe(results[['Description', 'Barcode', 'Preço Tabela', 'Preço c/ Desc']], use_container_width=True)
                
                # Seleção para Pedido
                st.divider()
                selected_prod = st.selectbox("Selecione o produto para o pedido:", results['Description'].tolist())
                prod_data = results[results['Description'] == selected_prod].iloc[0]
                
                c_v1, c_v2, c_v3 = st.columns(3)
                qtd = c_v1.number_input("Quantidade", 1, 1000, 1)
                desc_ind = c_v2.number_input("Desconto Unitário (%)", 0.0, 100.0, desc_geral)
                
                preco_venda = prod_data['Price'] * (1 - (desc_ind/100))
                c_v3.metric("Preço Unit. Venda", f"R$ {extra_round(preco_venda)}")
                
                if st.button("➕ Adicionar ao Pedido"):
                    item = {
                        "Descrição": prod_data['Description'],
                        "EAN": prod_data['Barcode'],
                        "Qtd": qtd,
                        "Preço Unit": extra_round(preco_venda),
                        "Total": extra_round(preco_venda * qtd)
                    }
                    st.session_state.carrinho.append(item)
                    st.toast("Item adicionado!")
            else:
                st.error("Nenhum produto encontrado.")

    with col_c2:
        st.subheader("🛒 Resumo do Pedido")
        if st.session_state.carrinho:
            df_cart = pd.DataFrame(st.session_state.carrinho)
            st.table(df_cart[['Descrição', 'Qtd', 'Total']])
            total_geral = df_cart['Total'].sum()
            st.metric("TOTAL DO PEDIDO", f"R$ {extra_round(total_geral)}")
            
            if st.button("🗑️ Limpar Pedido"):
                st.session_state.carrinho = []
                st.rerun()
            
            # Exportar Pedido
            csv_pedido = df_cart.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Baixar Pré-Pedido (CSV)", csv_pedido, f"pedido_{datetime.now().strftime('%H%M%S')}.csv")
        else:
            st.info("O carrinho está vazio.")

# --- ABA 3: GERENCIAR BANCO (MANTIDA) ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco")
    f = st.file_uploader("Upload Banco", type=["xlsx", "csv"])
    if f and st.button("💾 Salvar Banco"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        df = df.iloc[:, [0, 1, 2]]
        df.columns = ['Description', 'Barcode', 'Price']
        df['Barcode'] = df['Barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
        df.to_csv(DB_STORAGE, index=False)
        st.cache_data.clear()
        st.success("Banco Atualizado!")

# --- ABA 4: USUÁRIOS (MANTIDA) ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    
    # Backup/Restauração
    c_b1, c_b2 = st.columns(2)
    with c_b1:
        df_u = pd.DataFrame.from_dict(users, orient='index').reset_index()
        df_u.columns = ['Usuario', 'Senha', 'Expiracao', 'Nivel']
        st.download_button("📥 Baixar Backup Usuários", df_u.to_csv(index=False).encode('utf-8'), "backup_users.csv")
    with c_b2:
        up_b = st.file_uploader("📤 Restaurar Usuários", type=["csv"])
        if up_b and st.button("🔥 Restaurar"):
            df_r = pd.read_csv(up_b)
            new_u = {str(r['Usuario']): {"password": str(r['Senha']), "expiry": str(r['Expiracao']), "role": str(r['Nivel'])} for _, r in df_r.iterrows()}
            save_users(new_u); st.success("Restaurado!"); st.rerun()

    st.divider()
    col_n, col_e = st.columns(2)
    with col_n:
        st.subheader("➕ Novo")
        with st.form("Novo"):
            nu, np, nd = st.text_input("Usuário"), st.text_input("Senha"), st.number_input("Validade (dias)", 1, 365, 30)
            if st.form_submit_button("Criar"):
                users[nu] = {"password": np, "expiry": (datetime.now()+timedelta(days=nd)).strftime("%Y-%m-%d"), "role": "user"}
                save_users(users); st.success("Criado!"); st.rerun()
    with col_e:
        st.subheader("📝 Editar")
        u_sel = st.selectbox("Usuário", list(users.keys()))
        if u_sel:
            ep = st.text_input("Senha", value=users[u_sel]["password"])
            ex = st.text_input("Expiração", value=users[u_sel]["expiry"])
            if st.button("Salvar"):
                users[u_sel].update({"password": ep, "expiry": ex})
                save_users(users); st.success("Ok!"); st.rerun()
            if u_sel != "admin" and st.button("❌ Excluir"):
                del users[u_sel]; save_users(users); st.rerun()
