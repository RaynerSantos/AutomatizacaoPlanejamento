from models.workbook import WorkbookManager

import streamlit as st
import os

st.title("Atualizar Projetos de Hoje")

# Upload do arquivo
uploaded_file = st.file_uploader("Faça o upload do arquivo")
if uploaded_file:
    filename = uploaded_file.name

    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    filename = os.path.join(parent_dir, filename)

    workbook = WorkbookManager(filename)

    # Pega os projetos de hoje
    projetos = workbook.planilhas_semanais.get_projetos_disponiveis()
    coletas_por_projeto = {}
    hc_por_projeto = {}

    workbook.close()

    st.write("### Projetos disponíveis para hoje:")
    with st.form(key='my_form'):
        for projeto in projetos:
            coleta = st.number_input(f'Digite a coleta para o projeto "{projeto}": ')
            coletas_por_projeto[projeto] = coleta
            hc = st.number_input(f'Digite o hc para o projeto "{projeto}": ')
            hc_por_projeto[projeto] = hc
            st.markdown("---")

        submit = st.form_submit_button("Preencher")

        if submit:
            workbook = WorkbookManager(filename)
            workbook.planilhas_semanais.atualizar_coleta_diaria(coletas_por_projeto, hc_por_projeto)
            workbook.planilhas_semanais.atualizar_meta_parcial()
            workbook.save()
            st.success("Planilha atualizada com sucesso!")
