"""Search component for public course search."""
import streamlit as st
from utils.styling import load_css
from utils.excel_handler import get_excel_dataframes, search_in_dataframes

def show_search_page():
    """Display the public search page."""
    load_css()
    
    st.title("🔍 Buscar Mis Cursos")
    st.markdown("""
    Ingresa tu nombre completo o parte de él para buscar tus cursos registrados.
    """)
    
    # Check if Excel file is loaded in session state
    if 'excel_data_path' not in st.session_state or st.session_state.excel_data_path is None:
        st.warning("⚠️ No hay archivo de datos cargado. Por favor, contacta al administrador.")
        if st.button("🏠 Volver al inicio"):
            st.session_state.view = 'landing'
            st.rerun()
        return
    
    # Load dataframes
    try:
        honorarios_df, activos_df, _, _ = get_excel_dataframes(
            st.session_state.excel_workbook,
            st.session_state.excel_data_path
        )
    except Exception as e:
        st.error(f"Error al cargar los datos: {str(e)}")
        return
    
    # Search input
    search_term = st.text_input(
        "Buscar por nombre:",
        placeholder="Ej: Juan Pérez",
        key="search_input"
    )
    
    if search_term:
        # Perform search
        results_honorarios, results_activos = search_in_dataframes(
            honorarios_df, activos_df, search_term
        )
        
        # Display results
        has_results = False
        
        if results_honorarios is not None and not results_honorarios.empty:
            has_results = True
            st.subheader("📋 Cursos en HONORARIOS")
            st.dataframe(results_honorarios, use_container_width=True)
        
        if results_activos is not None and not results_activos.empty:
            has_results = True
            st.subheader("📋 Cursos en ACTIVOS")
            st.dataframe(results_activos, use_container_width=True)
        
        if not has_results:
            st.info("No se encontraron resultados para tu búsqueda.")
    else:
        st.info("👆 Ingresa tu nombre en el campo de búsqueda para ver tus cursos.")
    
    # Back button
    st.divider()
    if st.button("🏠 Volver al inicio"):
        st.session_state.view = 'landing'
        st.rerun()

