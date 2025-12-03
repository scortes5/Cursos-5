"""Authentication utilities."""
import os
from dotenv import load_dotenv

load_dotenv()

EDITOR_USERNAME = os.getenv('EDITOR_USERNAME')
EDITOR_PASSWORD = os.getenv('EDITOR_PASSWORD')

def is_authenticated(username, password):
    """Check if credentials are valid."""
    return username == EDITOR_USERNAME and password == EDITOR_PASSWORD

def require_auth():
    """Check if user is authenticated in session state."""
    import streamlit as st
    return st.session_state.get('authenticated', False)

