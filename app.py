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

# --- ABA 1: COTAÇÃO (CORREÇÃO DEFINITIVA DO CONTADOR) ---
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
            
            # Mapeia índices de coluna
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            try:
                d_idx = col_indices[desc_col.strip()]
                b_idx = col_indices[bar_col.strip()]
                p_idx = col_indices[price_col.strip()]
                
                count = 0
                
                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = ws.cell(row=r, column=b_idx).value
                    
                    # Se não tem descrição, ignoramos a linha completamente (evita rodapés vazios)
                    if d_val is None or str(d_val).strip() == "":
                        continue
                    
                    found_p = None
                    
                    # 1. Busca por Barras
                    if "Barras" in modo or "Híbrido" in modo:
                        barcodes = extract_all_barcodes(b_val)
                        for b in barcodes:
                            if b in price_map:
                                found_p = price_map[b]
                                break
                    
                    # 2. Busca por Similaridade
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo):
                        best_sim = 0
                        d_det = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim = similarity(d_val, row_db['Description'])
                            if sim >= 0.75 and d_det == extrair_detalhes(row_db['Description']):
                                if sim > best_sim:
                                    best_sim = sim
                                    found_p = row_db['Price']
                    
                    # 3. Preenchimento e Contagem real
                    if found_p is not None:
                        try:
                            # Calcula valor final
                            f_p = float(found_p) * (1 - (discount/100))
                            valor_final = extra_round(f_p) if aplicar_arredondamento else f_p
                            
                            # Grava na planilha
                            ws.cell(row=r, column=p_idx).value = valor_final
                            
                            # SÓ incrementa se o valor gravado for um número (confirmação final)
                            if isinstance(ws.cell(row=r, column=p_idx).value, (int, float)):
                                count += 1
                        except:
                            continue

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Processamento concluído: {count} itens preenchidos com preços.")
                st.download_button("📥 Baixar Planilha", out.getvalue(), "cotacao_final.xlsx")
                
            except KeyError:
                st.error("Erro: Colunas selecionadas não encontradas na linha de cabeçalho.")
            except Exception as e:
                st.error(f"Erro inesperado: {e}")

# --- ABA 2: VENDAS (MANTIDA) ---
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
                        st.write(f"Deseja replicar Qtd: {rep['qtd']} e Preço: R$ {rep['preco']} para família '{rep['familia']}'?")
                        familia_results = results[results['Description'].str.contains(rep['familia'], case=False) & (results['Barcode'] != rep['ean'])]
                        escolha = st.radio("Como deseja replicar?", ["Replicar em todos", "Escolher específicos"], horizontal=True)
                        
                        confirmar = False
                        if escolha == "Escolher específicos":
                            selecionados = st.multiselect("Itens:", familia_results['Description'].tolist(), default=familia_results['Description'].tolist())
                            if st.button("Confirmar Seleção"):
                                for _, rf in familia_results.iterrows():
                                    if rf['Description'] in selecionados:
                                        st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                                confirmar = True
                        else:
                            if st.button("Confirmar em Todos"):
                                for _, rf in familia_results.iterrows():
                                    st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                                confirmar = True
                        if confirmar:
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
            else:
                st.error("Nenhum produto encontrado.")

    with col_vendas_2:
        st.subheader("🛒 Seu Pedido")
        if st.session_state.carrinho:
            df_cart = pd.DataFrame(st.session_state.carrinho)
            for item in st.session_state.carrinho:
                st.write(f"{item['Qtd']}x {item['Descrição']} - **R$ {item['Total']}**")
            total_pedido = df_cart['Total'].sum()
            st.metric("Total Geral", f"R$ {extra_round(total_pedido)}")
            if st.button("🗑️ Limpar Tudo"):
                st.session_state.carrinho = []
                st.rerun()
            output_xlsx = io.BytesIO()
            wb_ped = openpyxl.Workbook()
            ws_ped = wb_ped.active
            headers = ["Descrição", "Código EAN", "Qtd", "Preço Unit.", "Total"]
            ws_ped.append(headers)
            for item in st.session_state.carrinho:
                ws_ped.append([item['Descrição'], item['EAN'], item['Qtd'], item['Preço Unit'], item['Total']])
            ws_ped.append(["", "", "", "TOTAL GERAL:", extra_round(total_pedido)])
            wb_ped.save(output_xlsx)
            st.download_button("📥 Baixar Pedido", output_xlsx.getvalue(), f"pedido_{datetime.now().strftime('%d_%m_%H%M%S')}.xlsx")
        else:
            st.info("Carrinho vazio.")

# --- ABA 3: GERENCIAR BANCO ---
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

# --- ABA 4: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    c_b1, c_b2 = st.columns(2)
    with c_b1:
        df_u = pd.DataFrame.from_dict(users, orient='index').reset_index()
        st.download_button("📥 Backup Usuários", df_u.to_csv(index=False).encode('utf-8'), "backup_users.csv")
    with c_b2:
        up_b = st.file_uploader("Restore Usuários", type=["csv"])
        if up_b and st.button("🔥 Restaurar"):
            df_r = pd.read_csv(up_b)
            new_u = {str(r['index']): {"password": str(r['password']), "expiry": str(r['expiry']), "role": str(r['role'])} for _, r in df_r.iterrows()}
            save_users(new_u); st.rerun()
    st.divider()
    col_n, col_e = st.columns(2)
    with col_n:
        st.subheader("➕ Novo")
        with st.form("Novo"):
            nu, np, nd = st.text_input("Usuário"), st.text_input("Senha"), st.number_input("Validade (dias)", 1, 365, 30)
            if st.form_submit_button("Criar"):
                users[nu] = {"password": np, "expiry": (datetime.now()+timedelta(days=nd)).strftime("%Y-%m-%d"), "role": "user"}
                save_users(users); st.rerun()
    with col_e:
        st.subheader("📝 Editar")
        u_sel = st.selectbox("Usuário", list(users.keys()))
        if u_sel:
            ep = st.text_input("Senha", value=users[u_sel]["password"])
            ex = st.text_input("Expiração", value=users[u_sel]["expiry"])
            if st.button("Salvar"):
                users[u_sel].update({"password": ep, "expiry": ex})
                save_users(users); st.rerun()
            if u_sel != "admin" and st.button("❌ Excluir"):
                del users[u_sel]; save_users(users); st.rerun()
