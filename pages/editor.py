"""Editor page for admin users."""
import streamlit as st
from datetime import datetime, date
import io
import tempfile
import os
from utils.styling import load_css
from utils.auth import require_auth
from utils.excel_handler import (
    load_excel_with_formatting,
    get_excel_dataframes,
    save_excel_with_formatting,
    get_sheet_columns,
    add_row_to_excel
)

def show_editor_page():
    """Display the editor page (admin only)."""
    if not require_auth():
        st.error("🔒 Acceso denegado. Debes iniciar sesión como administrador.")
        if st.button("Volver al inicio"):
            st.session_state.view = 'landing'
            st.rerun()
        return
    
    load_css()
    st.title("✏️ Modo Editor (Admin)")
    st.write("Aquí puedes editar los registros de cursos.")
    
    # Initialize session state for workbook
    if 'workbook' not in st.session_state:
        st.session_state.workbook = None
    if 'excel_file_name' not in st.session_state:
        st.session_state.excel_file_name = None
    if 'honorarios_sheet' not in st.session_state:
        st.session_state.honorarios_sheet = None
    if 'activos_sheet' not in st.session_state:
        st.session_state.activos_sheet = None
    
    # File uploader (only for admin)
    st.subheader("📤 Subir Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona un archivo Excel",
        type=['xlsx', 'xls'],
        help="El archivo debe contener hojas llamadas 'HONORARIOS' y 'ACTIVOS'"
    )
    
    if uploaded_file is not None:
        # Save uploaded file temporarily
        if st.session_state.excel_file_name != uploaded_file.name or st.session_state.workbook is None:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            # Load workbook with formatting
            workbook = load_excel_with_formatting(tmp_path)
            if workbook:
                st.session_state.workbook = workbook
                st.session_state.excel_file_name = uploaded_file.name
                st.session_state.tmp_path = tmp_path
                
                # Store in global session state for search functionality
                st.session_state.excel_data_path = tmp_path
                st.session_state.excel_workbook = workbook
                
                # Get sheet names
                sheet_names = workbook.sheetnames
                st.session_state.honorarios_sheet = next((name for name in sheet_names if name.upper() == 'HONORARIOS'), None)
                st.session_state.activos_sheet = next((name for name in sheet_names if name.upper() == 'ACTIVOS'), None)
                
                st.success(f"✅ Archivo cargado: {uploaded_file.name}")
                st.info(f"📋 Hojas encontradas: {', '.join(sheet_names)}")
    
    # Show form if workbook is loaded
    if st.session_state.workbook is not None:
        st.divider()
        st.subheader("📝 Formulario de Registro")
        
        # Sheet selection
        sheet_option = st.radio(
            "Selecciona la hoja donde agregar el registro:",
            ["HONORARIOS", "ACTIVOS"],
            horizontal=True
        )
        
        # Get target sheet name
        target_sheet = st.session_state.honorarios_sheet if sheet_option == "HONORARIOS" else st.session_state.activos_sheet
        
        if target_sheet:
            # Get columns from the selected sheet
            columns = get_sheet_columns(st.session_state.workbook, target_sheet)
            
            if columns:
                # Create form fields dynamically based on columns
                form_data = {}
                column_mapping = {}  # Maps column name to column index
                
                # Create mapping
                for idx, col_name in enumerate(columns, 1):
                    column_mapping[col_name] = idx
                
                # Organize form in columns
                num_cols = len(columns)
                cols_per_row = 2
                num_rows = (num_cols + cols_per_row - 1) // cols_per_row
                
                for row_idx in range(num_rows):
                    cols = st.columns(cols_per_row)
                    for col_idx in range(cols_per_row):
                        field_idx = row_idx * cols_per_row + col_idx
                        if field_idx < len(columns):
                            col_name = columns[field_idx]
                            with cols[col_idx]:
                                # Determine input type based on column name
                                col_lower = col_name.lower()
                                if 'fecha' in col_lower or 'date' in col_lower:
                                    date_val = st.date_input(
                                        col_name,
                                        value=datetime.now().date(),
                                        key=f"form_{target_sheet}_{col_name}"
                                    )
                                    form_data[col_name] = date_val
                                elif 'observacion' in col_lower or 'nota' in col_lower or 'comentario' in col_lower:
                                    form_data[col_name] = st.text_area(
                                        col_name,
                                        placeholder=f"Ingresa {col_name.lower()}",
                                        key=f"form_{target_sheet}_{col_name}"
                                    )
                                else:
                                    form_data[col_name] = st.text_input(
                                        col_name,
                                        placeholder=f"Ingresa {col_name.lower()}",
                                        key=f"form_{target_sheet}_{col_name}"
                                    )
                
                # Submit button
                st.markdown("---")
                col_submit1, col_submit2, col_submit3 = st.columns([1, 2, 1])
                with col_submit2:
                    if st.button("➕ Agregar Registro", type="primary", use_container_width=True):
                        # Prepare row data
                        row_data = {}
                        for col_name, value in form_data.items():
                            if isinstance(value, (datetime, date)):
                                row_data[col_name] = value.strftime("%d/%m/%Y")
                            else:
                                row_data[col_name] = str(value) if value else ""
                        
                        # Add row to workbook
                        if add_row_to_excel(st.session_state.workbook, target_sheet, row_data, column_mapping):
                            st.success(f"✅ Registro agregado exitosamente a la hoja {sheet_option}")
                            # Save the workbook
                            save_excel_with_formatting(st.session_state.workbook, st.session_state.tmp_path)
                            # Update global workbook reference
                            st.session_state.excel_workbook = st.session_state.workbook
                            # Clear form by rerunning
                            st.rerun()
                        else:
                            st.error("❌ Error al agregar el registro")
            else:
                st.warning("⚠️ No se pudieron detectar las columnas del archivo Excel")
        else:
            st.error(f"❌ No se encontró la hoja {sheet_option}")
        
        st.divider()
        
        # Download button
        st.subheader("💾 Descargar Archivo Actualizado")
        if st.button("⬇️ Descargar Excel Actualizado", type="primary"):
            if st.session_state.workbook:
                # Save workbook to bytes
                output = io.BytesIO()
                st.session_state.workbook.save(output)
                output.seek(0)
                
                st.download_button(
                    label="📥 Descargar",
                    data=output,
                    file_name=f"actualizado_{st.session_state.excel_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        # Preview current data
        st.divider()
        st.subheader("👀 Vista Previa de Datos")
        
        honorarios_df, activos_df, _, _ = get_excel_dataframes(st.session_state.workbook, st.session_state.tmp_path)
        
        tab1, tab2 = st.tabs(["HONORARIOS", "ACTIVOS"])
        
        with tab1:
            if honorarios_df is not None:
                st.dataframe(honorarios_df, use_container_width=True)
            else:
                st.info("No hay datos en la hoja HONORARIOS")
        
        with tab2:
            if activos_df is not None:
                st.dataframe(activos_df, use_container_width=True)
            else:
                st.info("No hay datos en la hoja ACTIVOS")
    
    # Logout button
    st.divider()
    if st.button("Cerrar sesión"):
        st.session_state.authenticated = False
        st.session_state.view = 'landing'
        # Clean up
        if 'workbook' in st.session_state:
            del st.session_state.workbook
        if 'tmp_path' in st.session_state:
            try:
                os.unlink(st.session_state.tmp_path)
            except:
                pass
        st.rerun()

