from models.projeto import Projeto
from models.workbook import WorkbookManager

import streamlit as st
import pythoncom
import math
import os

st.title("Adicionar Projeto")

# Informações básicas
uploaded_file = st.file_uploader("Faça o upload do arquivo")
if uploaded_file:
    filename = uploaded_file.name

    nome = st.text_input("Nome do projeto:")
    data_inicio = st.date_input("Data de início:", format="DD/MM/YYYY", value=None)
    duracao = st.number_input("Duração da coleta (dias úteis):", step=1, min_value=0)
    amostra = st.number_input("Amostra:", step=1, min_value=0)

    if nome and data_inicio and duracao and amostra:
        # Cria variável do projeto
        projeto = Projeto(nome=nome, amostra=amostra, data_inicio_coleta=data_inicio, duracao_coleta=duracao)

        # Informações de HC e produtividade
        media = st.number_input("Digite a média de entrevistas por dia:", step=1, min_value=0)
        if media:
            hc_sugerido = math.ceil(amostra / duracao / media)

            entrevistas_total = 0
            num_semana = 1
            for semana in projeto.semanas:
                entrevistas_semana = 0
                st.write("### Semana do dia " + semana.data_inicio.strftime("%d/%m/%Y"))
                
                # Criar colunas para cada dia da semana
                colunas = st.columns(len(semana.dias_uteis))

                for i, dia in enumerate(semana.dias_uteis):
                    if dia.data >= projeto.data_inicio_coleta and dia.data <= projeto.data_fim:
                        with colunas[i]:
                            hc = st.number_input("HC - " + dia.data.strftime("%d/%m/%Y"), step=1, value=hc_sugerido, min_value=0)
                            dia.hc = hc
                            entrevistas_semana += hc

                produtividade = st.number_input(f"Produtividade da semana {num_semana}", step=0.1, min_value=0.0)
                semana.produtividade = produtividade
                entrevistas_semana *= produtividade
                entrevistas_semana = round(entrevistas_semana)
                entrevistas_total += entrevistas_semana
                num_semana += 1

            gap = entrevistas_total - amostra
            st.write("Gap: " + str(gap))

            # Adicionando projeto
            if st.button("Adicionar Projeto"):
                pythoncom.CoInitialize()

                script_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(script_dir)
                filename = os.path.join(parent_dir, filename)

                num_dias_qt = 1
                num_dias_d = 3
                num_dias_h = 1
                num_dias_t = 1
                num_dias_c = duracao - (num_dias_t)
                specs = [
                    ("QT", num_dias_qt),
                    ("D", num_dias_d),   
                    ("H", num_dias_h),   
                    ("T", num_dias_t),   
                    ("C", num_dias_c),   
                ]

                workbook = WorkbookManager(filename)

                try:
                    workbook.cronograma_geral.inserir_projeto(specs, projeto)
                    workbook.proximas_semanas.inserir_projeto(projeto)
                    workbook.planilhas_semanais.inserir_projeto(projeto)
                    workbook.save()
                    st.write("Projeto adicionado!")
                except Exception as e:
                    st.error(f"Ocorreu um erro: {e}")
                finally:
                    workbook.close()
