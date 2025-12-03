"""Landing page."""
import streamlit as st
from utils.styling import load_css

def show_landing_page():
    """Display the landing page."""
    load_css()
    st.title("Cursos Quinta")
    st.subheader("""Agradecimientos:
                Sergio Cortés Prado
    Clemente Garmendia Pascal""")
    st.image("logo.png", width=70)
    st.markdown("""
    ## Bienvenido a la plataforma de cursos Quinta
    Aquí puedes gestionar y visualizar el registro de cursos de los voluntarios.\n
    - **🔍 Buscar**: Busca tus cursos ingresando tu nombre (sin necesidad de registro).\n
    - **✏️ Editor**: Acceso para modificar y actualizar los registros (requiere contraseña de administrador).\n
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔍 Buscar Mis Cursos", type="primary", use_container_width=True):
            st.session_state.view = 'search'
            st.rerun()
    with col2:
        if st.button("✏️ Editor (Admin)", use_container_width=True):
            st.session_state.view = 'editor_login'
            st.rerun()

