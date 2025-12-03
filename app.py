"""Main application entry point."""
import streamlit as st
from pages import landing, login, editor
from components import search

def main():
    """Main application router."""
    # Initialize session state
    if 'view' not in st.session_state:
        st.session_state.view = 'landing'
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'excel_data_path' not in st.session_state:
        st.session_state.excel_data_path = None
    if 'excel_workbook' not in st.session_state:
        st.session_state.excel_workbook = None
    
    # Route to appropriate page
    if st.session_state.view == 'landing':
        landing.show_landing_page()
    elif st.session_state.view == 'editor_login':
        login.show_login_page()
    elif st.session_state.view == 'editor':
        editor.show_editor_page()
    elif st.session_state.view == 'search':
        search.show_search_page()

if __name__ == "__main__":
    main()
