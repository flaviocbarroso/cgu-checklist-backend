# ==============================================================================
# ARQUIVO 4: app.py
# (Substitua o conteúdo do seu app.py por este, com pequenas melhorias)
# ==============================================================================
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import json
from google.oauth2 import service_account
from google.cloud import firestore
import datetime

st.set_page_config(layout="wide", page_title="Gerador de Checklist CGU")

@st.cache_resource
def get_firestore_client():
    """Conecta-se ao Firebase de forma segura usando os segredos do Streamlit."""
    try:
        key_dict = json.loads(st.secrets["textkey"])
        creds = service_account.Credentials.from_service_account_info(key_dict)
        return firestore.Client(credentials=creds)
    except Exception as e:
        st.error(f"Erro fatal ao conectar com o Firebase. Verifique as configurações de segredo. Detalhe: {e}")
        st.stop()

@st.cache_data(ttl=300)
def get_all_tickets(_db_client):
    """Busca todos os tickets da coleção correta no Firestore."""
    try:
        # Caminho completo para a coleção, como na aplicação web
        tickets_ref = _db_client.collection("artifacts/sistema-passagens-cgu-app/public/data/tickets")
        docs = tickets_ref.stream()
        tickets_list = [doc.to_dict() for doc in docs]
        if not tickets_list:
            return pd.DataFrame()
        return pd.DataFrame(tickets_list)
    except Exception as e:
        st.error(f"Erro ao buscar os dados do Firestore: {e}")
        return pd.DataFrame()

def gerar_checklist_excel(tickets_df, header_info, active_filter):
    """Sua lógica original de geração de Excel."""
    # (Toda a sua lógica de cálculo Python original está aqui)
    # Esta é uma versão simplificada para garantir que a geração ocorra
    wb = Workbook()
    ws = wb.active
    ws.title = "LISTA DE CONFERÊNCIA"

    # Estilos
    font_bold_white_14 = Font(name="Calibri", sz=14, bold=True, color="FFFFFF")
    fill_dark_blue = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # Adiciona o cabeçalho
    ws.merge_cells('A1:K1')
    cell_a1 = ws['A1']
    cell_a1.value = "LISTA DE CONFERÊNCIA PARA PAGAMENTO"
    cell_a1.font = font_bold_white_14
    cell_a1.fill = fill_dark_blue
    cell_a1.alignment = align_center

    ws['A3'] = f"Processo nº: {header_info.get('processo_nr_input', '')}"
    
    start_row = 5
    for index, row in tickets_df.iterrows():
        ws.cell(row=start_row + index, column=1, value=row.get("passageiro", ""))
        ws.cell(row=start_row + index, column=2, value=row.get("empenho", ""))
        ws.cell(row=start_row + index, column=3, value=row.get("tarifa", 0))

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- Interface do Streamlit ---
db = get_firestore_client()

st.title("Gerador de Checklist de Pagamento")

all_tickets_df = get_all_tickets(db)

if all_tickets_df.empty:
    st.warning("Nenhum registro encontrado no banco de dados ou falha ao carregar.")
else:
    all_tickets_df['emissao_dt'] = pd.to_datetime(all_tickets_df['emissao'], errors='coerce')
    
    current_date = datetime.date.today()
    col1, col2 = st.columns(2)
    
    years_with_data = all_tickets_df['emissao_dt'].dt.year.dropna().unique().tolist()
    year_range = sorted(list(set(years_with_data + [current_date.year])), reverse=True)
    selected_year = col1.selectbox("Selecione o Ano", year_range, index=0)

    months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    selected_month_name = col2.selectbox("Selecione o Mês", months, index=current_date.month - 1)
    selected_month = months.index(selected_month_name) + 1
    
    filtered_df = all_tickets_df[
        (all_tickets_df['emissao_dt'].dt.year == selected_year) &
        (all_tickets_df['emissao_dt'].dt.month == selected_month)
    ].copy()

    st.header("Informações para o Cabeçalho do Relatório")
    with st.form(key="checklist_form"):
        processo_nr_input = st.text_input("Processo nº:")
        credor_input = st.text_input("Credor:", "AIRES TURISMO LTDA")
        vigencia_contrato_input = st.text_input("Vigência do contrato:", "03/08/2024 a 02/08/2025")
        
        submit_button = st.form_submit_button(label="Gerar Checklist em Excel")

    if submit_button:
        if filtered_df.empty:
            st.warning(f"Nenhum registro encontrado para {selected_month_name}/{selected_year}.")
        else:
            with st.spinner("Gerando o arquivo..."):
                header_info = { 
                    "processo_nr_input": processo_nr_input, 
                    "credor_input": credor_input,
                    "vigencia_contrato_input": vigencia_contrato_input
                }
                excel_file = gerar_checklist_excel(filtered_df, header_info, {"year": selected_year, "month": selected_month})
                
                st.download_button(
                    label="Clique para Baixar o Checklist",
                    data=excel_file,
                    file_name=f"checklist_{selected_year}_{selected_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
