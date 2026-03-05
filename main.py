import streamlit as st
import os

st.set_page_config(page_title="Kaufmann CRM", layout="wide")
LOGO_PATH = "Logo_Kaufmann.jpg"

USUARIOS = {
    "adm_kaufmann": {"senha": "123", "perfil": "ADM", "nome": "Administrador"},
    "vendedor_01": {"senha": "789", "perfil": "VEND", "nome": "Vendedor 1"},
}

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, use_container_width=True)
        st.markdown("<h2 style='text-align: center;'>Acesso ao Sistema</h2>", unsafe_allow_html=True)
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        if st.button("ENTRAR", use_container_width=True):
            if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
                st.session_state.logged_in = True
                st.session_state.user_data = USUARIOS[usuario]
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
else:
    st.sidebar.success(f"Logado como: {st.session_state.user_data['nome']}")
    st.title("🚀 Bem-vindo ao CRM Kaufmann")
    st.write("Selecione uma opção no menu ao lado para começar.")
    if st.sidebar.button("Sair"):
        st.session_state.logged_in = False
        st.rerun()