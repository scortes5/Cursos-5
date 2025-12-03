"""Login page for admin authentication."""
import streamlit as st
from utils.styling import load_css
from utils.auth import is_authenticated

def show_login_page():
    """Display the login page."""
    load_css()
    st.title("🔐 Acceso Editor")
    st.write("Por favor, ingresa tus credenciales para acceder al modo editor.")
    
    username = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Ingresar", type="primary", use_container_width=True):
            if is_authenticated(username, password):
                st.session_state.authenticated = True
                st.session_state.view = 'editor'
                st.success("✅ Acceso concedido")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    
    with col2:
        if st.button("Volver", use_container_width=True):
            st.session_state.view = 'landing'
            st.rerun()

