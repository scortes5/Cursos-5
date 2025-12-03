"""Excel file handling utilities with formatting preservation."""
import streamlit as st
import pandas as pd
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border
import io

def load_excel_with_formatting(file_path):
    """Load Excel file using openpyxl to preserve formatting."""
    try:
        workbook = load_workbook(file_path, data_only=False)
        return workbook
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None

def get_excel_dataframes(workbook, tmp_path):
    """Extract data from workbook as dataframes for easier manipulation."""
    try:
        # Find sheets by case-insensitive matching
        sheet_names = workbook.sheetnames
        honorarios_sheet = next((name for name in sheet_names if name.upper() == 'HONORARIOS'), None)
        activos_sheet = next((name for name in sheet_names if name.upper() == 'ACTIVOS'), None)
        
        if not honorarios_sheet or not activos_sheet:
            return None, None, None, None
        
        # Save workbook temporarily to read with pandas
        workbook.save(tmp_path)
        
        # Convert sheets to dataframes
        honorarios_df = pd.read_excel(tmp_path, sheet_name=honorarios_sheet)
        activos_df = pd.read_excel(tmp_path, sheet_name=activos_sheet)
        
        return honorarios_df, activos_df, honorarios_sheet, activos_sheet
    except Exception as e:
        return None, None, None, None

def save_excel_with_formatting(workbook, file_path):
    """Save workbook maintaining all formatting."""
    try:
        workbook.save(file_path)
        return True
    except Exception as e:
        st.error(f"Error saving Excel file: {str(e)}")
        return False

def get_sheet_columns(workbook, sheet_name):
    """Get column names from the Excel sheet."""
    try:
        sheet = workbook[sheet_name]
        columns = []
        
        # Try to find header row (usually row 1, but check first few rows)
        for row_idx in range(1, min(5, sheet.max_row + 1)):
            row_values = [cell.value for cell in sheet[row_idx]]
            if any(val and str(val).strip() for val in row_values):
                columns = [str(val).strip() if val else f"Columna {i+1}" for i, val in enumerate(row_values)]
                break
        
        return columns
    except Exception as e:
        return []

def add_row_to_excel(workbook, sheet_name, row_data, column_mapping=None):
    """Add a new row to the specified sheet while preserving formatting."""
    try:
        sheet = workbook[sheet_name]
        
        # Find the next empty row (skip header if exists)
        data_start_row = 1
        for row in range(1, min(sheet.max_row + 1, 10)):
            if any(cell.value for cell in sheet[row]):
                data_start_row = row
                break
        
        # Find the last row with data
        last_data_row = sheet.max_row
        for row in range(sheet.max_row, 0, -1):
            if any(cell.value for cell in sheet[row]):
                last_data_row = row
                break
        
        next_row = last_data_row + 1
        
        # If column_mapping is provided, use it to map values to columns
        if column_mapping:
            col_data = {}
            for col_name, value in row_data.items():
                if col_name in column_mapping:
                    col_idx = column_mapping[col_name]
                    col_data[col_idx] = value
            
            # Write values to correct columns
            for col_idx, value in col_data.items():
                cell = sheet.cell(row=next_row, column=col_idx)
                cell.value = value
                
                # Copy formatting from last data row
                if last_data_row >= data_start_row:
                    source_cell = sheet.cell(row=last_data_row, column=col_idx)
                    if source_cell.font:
                        cell.font = Font(
                            name=source_cell.font.name if source_cell.font.name else 'Calibri',
                            size=source_cell.font.size if source_cell.font.size else 11,
                            bold=source_cell.font.bold,
                            italic=source_cell.font.italic,
                            color=source_cell.font.color
                        )
                    if source_cell.alignment:
                        cell.alignment = Alignment(
                            horizontal=source_cell.alignment.horizontal if source_cell.alignment.horizontal else 'general',
                            vertical=source_cell.alignment.vertical if source_cell.alignment.vertical else 'bottom',
                            wrap_text=source_cell.alignment.wrap_text if source_cell.alignment.wrap_text else False
                        )
                    if source_cell.fill and source_cell.fill.fill_type:
                        cell.fill = PatternFill(
                            fill_type=source_cell.fill.fill_type,
                            start_color=source_cell.fill.start_color,
                            end_color=source_cell.fill.end_color
                        )
                    if source_cell.border:
                        cell.border = Border(
                            left=source_cell.border.left,
                            right=source_cell.border.right,
                            top=source_cell.border.top,
                            bottom=source_cell.border.bottom
                        )
        else:
            # Fallback: write in order
            source_row = sheet[last_data_row] if last_data_row >= data_start_row else None
            for col_idx, value in enumerate(row_data.values() if isinstance(row_data, dict) else row_data, 1):
                cell = sheet.cell(row=next_row, column=col_idx)
                cell.value = value
                
                if source_row and col_idx <= len(source_row):
                    source_cell = source_row[col_idx - 1]
                    if source_cell.font:
                        cell.font = Font(
                            name=source_cell.font.name if source_cell.font.name else 'Calibri',
                            size=source_cell.font.size if source_cell.font.size else 11,
                            bold=source_cell.font.bold,
                            italic=source_cell.font.italic,
                            color=source_cell.font.color
                        )
                    if source_cell.alignment:
                        cell.alignment = Alignment(
                            horizontal=source_cell.alignment.horizontal if source_cell.alignment.horizontal else 'general',
                            vertical=source_cell.alignment.vertical if source_cell.alignment.vertical else 'bottom',
                            wrap_text=source_cell.alignment.wrap_text if source_cell.alignment.wrap_text else False
                        )
        
        return True
    except Exception as e:
        st.error(f"Error adding row: {str(e)}")
        return False

def search_in_dataframes(honorarios_df, activos_df, search_term):
    """Search for a name in both dataframes."""
    if honorarios_df is None and activos_df is None:
        return None, None
    
    results_honorarios = None
    results_activos = None
    
    # Normalize search term
    search_term = str(search_term).strip().lower()
    
    if honorarios_df is not None:
        # Search in all string columns
        mask = honorarios_df.astype(str).apply(
            lambda x: x.str.lower().str.contains(search_term, na=False)
        ).any(axis=1)
        results_honorarios = honorarios_df[mask] if mask.any() else None
    
    if activos_df is not None:
        # Search in all string columns
        mask = activos_df.astype(str).apply(
            lambda x: x.str.lower().str.contains(search_term, na=False)
        ).any(axis=1)
        results_activos = activos_df[mask] if mask.any() else None
    
    return results_honorarios, results_activos

