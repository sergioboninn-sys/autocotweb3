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

if not os.path.exists(USERS_STORAGE):
    with open(USERS_STORAGE, "w") as f:
        json.dump({"admin": {"password": "admin123", "expiry": "2099-12-31", "role": "admin"}}, f)

def load_users():
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

def try_read_excel(file):
    """Tenta ler Excel de várias formas (XLSX, XLS antigo, XLS-HTML)"""
    file.seek(0)
    try:
        # Tenta padrão (XLSX)
        return pd.read_excel(file)
    except:
        try:
            # Tenta XLS antigo
            file.seek(0)
            return pd.read_excel(file, engine='xlrd')
        except:
            try:
                # Tenta se for um HTML disfarçado de XLS (comum em sistemas antigos)
                file.seek(0)
                return pd.read_html(file)[0]
            except:
                file.seek(0)
                # Tenta binário
                return pd.read_excel(file, engine='pyxlsb')

# --- CARREGAMENTO DO BANCO ---
@st.cache_data
def get_master_db():
    if os.path.exists(DB_STORAGE):
        df = pd.read_csv(DB_STORAGE)
        df['Barcode'] = df['Barcode'].astype(str).str.replace(r'\.0$', '', regex=True)
        return df
    return pd.DataFrame(columns=['Description', 'Barcode', 'Price'])

# --- INTERFACE ---
tabs = ["📊 Cotação", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    if master_db.empty:
        st.warning("O banco de dados está vazio. Vá em 'Gerenciar Banco' primeiro.")
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido (Barras + Similaridade)", "Apenas Barras", "Apenas Similaridade"])
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Planilha de Destino", type=["xlsx", "xls", "xlsb"])

    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início dos produtos:", 1, 1000, 11)

        try:
            # Leitura robusta para a visualização
            t_df_view = try_read_excel(target_file)
            # Ajusta o header na visualização conforme input do usuário
            t_df_view.columns = t_df_view.iloc[int(header_pos)-1]
            t_df_view = t_df_view.iloc[int(header_pos):]
            
            st.subheader("Mapeamento de Colunas")
            col1, col2, col3 = st.columns(3)
            desc_col = col1.selectbox("Coluna Descrição", t_df_view.columns)
            bar_col = col2.selectbox("Coluna Barras", t_df_view.columns)
            price_col = col3.selectbox("Coluna Preço", t_df_view.columns)

            if st.button("🚀 Processar e Gerar Nova Planilha"):
                price_map = dict(zip(master_db['Barcode'].astype(str), master_db['Price']))
                
                # Para arquivos problemáticos (.xls BIFF21), a melhor forma de garantir o sucesso
                # é reconstruir a planilha em um novo XLSX, pois o openpyxl não suporta arquivos corrompidos.
                
                new_data = try_read_excel(target_file)
                wb = openpyxl.Workbook()
                ws = wb.active
                
                # Copia dados para o novo workbook
                for r_idx, row in enumerate(new_data.values, 1):
                    for c_idx, value in enumerate(row, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)
                
                # Identifica índices das colunas (base 1)
                cols_list = list(new_data.columns)
                try:
                    d_idx = cols_list.index(desc_col) + 1
                    b_idx = cols_list.index(bar_col) + 1
                    p_idx = cols_list.index(price_col) + 1
                except:
                    st.error("Erro ao localizar colunas. Tente ajustar a 'Linha do Cabeçalho'.")
                    st.stop()

                count = 0
                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = ws.cell(row=r, column=b_idx).value
                    
                    found_p = None
                    if ("Barras" in modo or "Híbrido" in modo) and b_val:
                        bcodes = extract_all_barcodes(b_val)
                        for b in bcodes:
                            if b in price_map:
                                found_p = price_map[b]
                                break
                    
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo) and d_val:
                        best_sim = 0
                        d_detalhes = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim_val = similarity(d_val, row_db['Description'])
                            if sim_val >= 0.75 and d_detalhes == extrair_detalhes(row_db['Description']):
                                if sim_val > best_sim:
                                    best_sim = sim_val
                                    found_p = row_db['Price']

                    if found_p is not None:
                        final_p = float(found_p) * (1 - (discount / 100))
                        if aplicar_arredondamento:
                            final_p = extra_round(final_p)
                        ws.cell(row=r, column=p_idx).value = final_p
                        count += 1

                output = io.BytesIO()
                wb.save(output)
                st.success(f"Sucesso! {count} itens processados.")
                st.download_button("📥 Baixar Planilha Pronta (XLSX)", output.getvalue(), "cotacao_final.xlsx")
                
        except Exception as e:
            st.error(f"O arquivo enviado tem um formato incompatível ou está corrompido. Tente salvar como 'Pasta de Trabalho do Excel (.xlsx)' antes de subir. Erro: {e}")

# --- ABA 2: GERENCIAR BANCO ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    f = st.file_uploader("Upload Banco (Referência)", type=["xlsx", "xls", "csv", "xlsb"])
    if f and st.button("💾 Salvar e Atualizar Banco"):
        try:
            if f.name.lower().endswith('.csv'):
                df = pd.read_csv(f)
            else:
                df = try_read_excel(f)
                
            df = df.iloc[:, [0, 1, 2]]
            df.columns = ['Description', 'Barcode', 'Price']
            df['Barcode'] = df['Barcode'].apply(lambda x: re.sub(r'\D', '', str(x).split('.')[0]))
            df.to_csv(DB_STORAGE, index=False)
            st.cache_data.clear()
            st.success("Banco de dados atualizado com sucesso!")
        except Exception as e:
            st.error(f"Erro ao carregar banco: {e}")

# --- ABA 3: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Administração de Usuários")
    users = load_users()
    col_novo, col_edit = st.columns(2)
    with col_novo:
        st.subheader("➕ Novo Usuário")
        with st.form("Novo"):
            new_u = st.text_input("Nome do Usuário")
            new_p = st.text_input("Senha")
            days = st.number_input("Dias de validade", 1, 365, 30)
            if st.form_submit_button("Criar"):
                if new_u and new_p:
                    users[new_u] = {
                        "password": new_p, 
                        "expiry": (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d"),
                        "role": "user"
                    }
                    save_users(users)
                    st.success(f"Usuário {new_u} criado!")
                    st.rerun()
    with col_edit:
        st.subheader("📝 Editar/Excluir")
        user_to_edit = st.selectbox("Selecione o usuário", list(users.keys()))
        if user_to_edit:
            edit_p = st.text_input("Nova Senha", value=users[user_to_edit]["password"])
            edit_exp = st.text_input("Expiração (AAAA-MM-DD)", value=users[user_to_edit]["expiry"])
            c_btn1, c_btn2 = st.columns(2)
            if c_btn1.button("Salvar Alterações"):
                users[user_to_edit]["password"] = edit_p
                users[user_to_edit]["expiry"] = edit_exp
                save_users(users)
                st.success("Atualizado!")
                st.rerun()
            if user_to_edit != "admin":
                if c_btn2.button("❌ Excluir Usuário", type="primary"):
                    del users[user_to_edit]
                    save_users(users)
                    st.warning("Usuário removido.")
                    st.rerun()
