"""Styling utilities for the application."""
import streamlit as st

STYLE_PATH = 'style.css'

def load_css():
    """Load and apply CSS styles."""
    try:
        with open(STYLE_PATH) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"CSS file not found: {STYLE_PATH}")

