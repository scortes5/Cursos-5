import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Load CSS
def load_css():
    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

def load_excel(file_path):
    """Load both sheets from the Excel file with flexible sheet name handling."""
    try:
        # Read all sheet names
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        
        # Find sheets by case-insensitive matching
        honorarios_sheet = next((name for name in sheet_names if name.upper() == 'HONORARIOS '), None)
        activos_sheet = next((name for name in sheet_names if name.upper() == 'ACTIVOS '), None)
        
        if not honorarios_sheet:
            st.error("Could not find a sheet named 'HONORARIOS' (case-insensitive). Please check your Excel file.")
            return None, None
        if not activos_sheet:
            st.error("Could not find a sheet named 'ACTIVOS' (case-insensitive). Please check your Excel file.")
            return None, None
            
        # Load the sheets using their actual names
        honorarios = pd.read_excel(file_path, sheet_name=honorarios_sheet)
        activos = pd.read_excel(file_path, sheet_name=activos_sheet)
        
        # Rename unnamed columns to proper names
        column_mapping = {
            'Unnamed: 3': 'Nombre',
            'Unnamed: 4': 'Primer Apellido',
            'Unnamed: 5': 'Segundo Apellido'
        }
        
        honorarios = honorarios.rename(columns=column_mapping)
        activos = activos.rename(columns=column_mapping)
        
        # Display sheet names for debugging
        st.info(f"Found sheets: {', '.join(sheet_names)}")
        
        return honorarios, activos
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None, None

def save_excel(file_path, honorarios_df, activos_df):
    """Save both sheets to the Excel file."""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            honorarios_df.to_excel(writer, sheet_name='HONORARIOS ', index=False)
            activos_df.to_excel(writer, sheet_name='ACTIVOS ', index=False)
        return True
    except Exception as e:
        st.error(f"Error saving Excel file: {str(e)}")
        return False

def main():
    # Load CSS
    load_css()
    
    # Create two columns for the title and logo
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.title("Cursos Quinta")
        st.subheader("""Agradecimientos:
        Sergio Cortés Prado
    Clemente Garmendia Pascal
        """)
    
    with col2:
        st.image("logo.png", width=70)
    
    # File uploader
    uploaded_file = st.file_uploader("Subir Archivo de Excel (.xlsx)", type=['xlsx'])
    
    if uploaded_file is not None:
        # Save the uploaded file temporarily
        temp_path = "temp_excel.xlsx"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Load the Excel file
        honorarios_df, activos_df = load_excel(temp_path)
        
        if honorarios_df is not None and activos_df is not None:
            # Clean up volunteer names for selection
            honorarios_df['Full Name'] = honorarios_df['Nombre'].fillna('') + ' ' + honorarios_df['Primer Apellido'].fillna('') + ' ' + honorarios_df['Segundo Apellido'].fillna('')
            activos_df['Full Name'] = activos_df['Nombre'].fillna('') + ' ' + activos_df['Primer Apellido'].fillna('') + ' ' + activos_df['Segundo Apellido'].fillna('')
            
            # Remove invalid names and clean up whitespace
            valid_honorarios_volunteers = [name.strip() for name in honorarios_df['Full Name'].dropna() 
                                         if name != 'Nombre Primer Apellido Segundo Apellido' and name.strip()]
            valid_activos_volunteers = [name.strip() for name in activos_df['Full Name'].dropna() 
                                      if name != 'Nombre Primer Apellido Segundo Apellido' and name.strip()]
            
            # Remove duplicates and sort
            valid_honorarios_volunteers = sorted(list(set(valid_honorarios_volunteers)))
            valid_activos_volunteers = sorted(list(set(valid_activos_volunteers)))
            
            # Volunteer type selection
            volunteer_type = st.radio("Seleccione tipo de voluntario", ["HONORARIOS", "ACTIVOS"])
            
            # Select the appropriate dataframe based on volunteer type
            current_df = honorarios_df if volunteer_type == "HONORARIOS" else activos_df
            
            # Display available volunteers with first and last names
            selected_volunteer_name = st.selectbox("Seleccione Voluntario", valid_honorarios_volunteers if volunteer_type == "HONORARIOS" else valid_activos_volunteers)
            
            # Select the course from available columns
            # Get course names from the first row of data (row 0) and their corresponding column names
            course_data = current_df.iloc[0, 15:]
            course_columns = course_data.dropna().index.tolist()
            course_names = course_data.dropna().values.tolist()
            
            # Create a mapping of course names to their column names
            course_mapping = dict(zip(course_names, course_columns))
            
            # Initialize session state for storing updates if it doesn't exist
            if 'updates' not in st.session_state:
                st.session_state.updates = []

            # Course selection and date input
            selected_course_name = st.selectbox("Elija Curso", course_names)
            selected_course_column = course_mapping[selected_course_name]
            completion_date = st.date_input("Fecha Curso Completado")
            
            # Add button to collect updates
            if st.button("Agregar Curso"):
                if selected_volunteer_name:
                    # Find the volunteer in the dataframe
                    if selected_volunteer_name in current_df['Full Name'].values:
                        # Store the update in session state
                        st.session_state.updates.append({
                            'volunteer': selected_volunteer_name,
                            'course': selected_course_column,
                            'date': completion_date
                        })
                        st.success(f"Curso {selected_course_name} agregado para {selected_volunteer_name}")
                    else:
                        st.error("Voluntario no encontrado en la categoría...")
                else:
                    st.error("Seleccione un voluntario")

            # Display current updates
            if st.session_state.updates:
                st.subheader("Cursos a actualizar:")
                for i, update in enumerate(st.session_state.updates):
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        # Get the course name from the mapping
                        course_name = next((name for name, col in course_mapping.items() if col == update['course']), update['course'])
                        st.write(f"Voluntario: {update['volunteer']}, Curso: {course_name}, Fecha: {update['date']}")
                    with col2:
                        st.markdown("""
                            <style>
                                div[data-testid="stButton"] button {
                                    background-color: #DC2626;
                                }
                                div[data-testid="stButton"] button:hover {
                                    background-color: #B91C1C;
                                }
                            </style>
                        """, unsafe_allow_html=True)
                        if st.button("Eliminar", key=f"delete_{i}"):
                            st.session_state.updates.pop(i)
                            st.rerun()
                
                # Add clear all button
                if st.button("Eliminar Todos los Cambios"):
                    st.session_state.updates = []
                    st.rerun()

            # Save all updates button
            if st.button("Guardar Todos los Cambios"):
                if st.session_state.updates:
                    # Apply all updates
                    for update in st.session_state.updates:
                        current_df.loc[current_df['Full Name'] == update['volunteer'], update['course']] = update['date']
                    
                    # Save the updated Excel file
                    if save_excel(temp_path, honorarios_df, activos_df):
                        st.success("¡Todos los cambios han sido guardados!")
                        # Clear the updates after successful save
                        st.session_state.updates = []
                        
                        # Add download button after successful save
                        with open(temp_path, 'rb') as f:
                            st.download_button(
                                label="Descargar Archivo Actualizado",
                                data=f,
                                file_name="Registro_Quinta_Cursos.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.warning("No hay cambios para guardar")

            # Display current data
            st.subheader("Información Actual")
            st.dataframe(current_df)
            
            # Clean up temporary file
            os.remove(temp_path)



main()
