# ==============================================================================
# ARQUIVO 2: app.py
# (Este é o seu script, adaptado para uma aplicação Streamlit)
# ==============================================================================
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from decimal import Decimal, getcontext
import io
import json
from google.oauth2 import service_account
from google.cloud import firestore
import datetime

# --- Configuração do Firebase ---
# Usa os segredos do Streamlit para conectar ao Firebase de forma segura
try:
    key_dict = json.loads(st.secrets["textkey"])
    creds = service_account.Credentials.from_service_account_info(key_dict)
    db = firestore.Client(credentials=creds)
except Exception as e:
    st.error(f"Erro ao conectar com o Firebase. Verifique as configurações de segredo. Detalhe: {e}")
    st.stop()
    
# --- Funções do seu Script Original (Adaptadas) ---
def get_all_tickets():
    """Busca todos os tickets do Firestore e retorna como um DataFrame."""
    tickets_ref = db.collection("tickets") # Assumindo que a coleção se chama 'tickets'
    docs = tickets_ref.stream()
    tickets_list = [doc.to_dict() for doc in docs]
    return pd.DataFrame(tickets_list)

def gerar_checklist_excel(tickets_df, header_info, active_filter):
    """Sua lógica original de geração de Excel."""
    # (Toda a sua lógica de cálculo e formatação de openpyxl vai aqui,
    # usando tickets_df e header_info como entrada)
    
    # Exemplo simplificado para garantir a funcionalidade básica:
    wb = Workbook()
    ws = wb.active
    ws.title = "LISTA DE CONFERÊNCIA"

    # --- Estilos (como no seu script) ---
    font_bold_white_14 = Font(name="Calibri", sz=14, bold=True, color="FFFFFF")
    fill_dark_blue = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # --- Adiciona o cabeçalho ---
    ws.merge_cells('A1:K1')
    ws['A1'] = "LISTA DE CONFERÊNCIA PARA PAGAMENTO"
    ws['A1'].font = font_bold_white_14
    ws['A1'].fill = fill_dark_blue
    ws['A1'].alignment = align_center

    # --- Adiciona dados ---
    ws['A3'] = f"Processo nº: {header_info['processo_nr_input']}"
    
    # Adicionar os dados dos tickets na planilha
    start_row = 5
    for index, row in tickets_df.iterrows():
        ws.cell(row=start_row + index, column=1, value=row.get("passageiro", ""))
        ws.cell(row=start_row + index, column=2, value=row.get("empenho", ""))
        # ... e assim por diante para todas as colunas
        
    # Salva o arquivo em memória para download
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- Interface do Streamlit ---
st.set_page_config(layout="wide", page_title="Gerador de Checklist CGU")
st.title("Gerador de Checklist de Pagamento")

# --- Filtros de Data ---
current_date = datetime.date.today()
col1, col2 = st.columns(2)
selected_year = col1.selectbox("Selecione o Ano", range(2023, current_date.year + 2), index=range(2023, current_date.year + 2).index(current_date.year))
months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
selected_month_name = col2.selectbox("Selecione o Mês", months, index=current_date.month - 1)
selected_month = months.index(selected_month_name) + 1

# --- Formulário para Informações do Cabeçalho ---
st.header("Informações para o Cabeçalho do Relatório")

with st.form(key="checklist_form"):
    processo_nr_input = st.text_input("Processo nº:")
    credor_input = st.text_input("Credor:", "AIRES TURISMO LTDA")
    # ... Adicionar todos os outros campos de input do seu modal aqui
    
    submit_button = st.form_submit_button(label="Gerar Checklist em Excel")

if submit_button:
    with st.spinner("Buscando dados e gerando o arquivo..."):
        all_tickets_df = get_all_tickets()
        
        if all_tickets_df.empty:
            st.warning("Nenhum registro encontrado no banco de dados.")
        else:
            # Filtrar DataFrame pelo período selecionado
            all_tickets_df['emissao_dt'] = pd.to_datetime(all_tickets_df['emissao'])
            filtered_df = all_tickets_df[
                (all_tickets_df['emissao_dt'].dt.year == selected_year) &
                (all_tickets_df['emissao_dt'].dt.month == selected_month)
            ]

            if filtered_df.empty:
                st.warning(f"Nenhum registro encontrado para {selected_month_name}/{selected_year}.")
            else:
                st.success(f"{len(filtered_df)} registros encontrados para o período.")
                
                header_info = {
                    "processo_nr_input": processo_nr_input,
                    "credor_input": credor_input,
                    # ... coletar os outros valores do formulário
                }
                
                excel_file = gerar_checklist_excel(filtered_df, header_info, {"year": selected_year, "month": selected_month})
                
                st.download_button(
                    label="Clique para Baixar o Checklist",
                    data=excel_file,
                    file_name=f"checklist_{selected_year}_{selected_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
```
