import streamlit as st

# Título centralizado
st.markdown("<h1 style='text-align: center;'>Planejamento</h1>", unsafe_allow_html=True)

# Criar duas colunas para centralizar os botões
col1, col2, col3 = st.columns([1, 1, 1])

# Adicionar os botões na coluna do meio
with col2:
    st.link_button(label="Adicionar Projeto", url="Adicionar_Projeto")
    st.link_button(label="Atualização Diária", url='Atualização_Diária')
