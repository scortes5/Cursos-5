# Course Management System

This application allows you to manage course information for volunteers and update Excel files with course completion dates.

## Features

- Upload Excel files with two sheets: "Honorarios" and "Activos"
- Select volunteer type (Honorarios or Activos)
- Choose from available courses
- Enter volunteer name and course completion date
- Automatically update Excel file with new information
- View current data in the application

## Setup

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

## Excel File Format

The Excel file should have two sheets:
1. "Honorarios" - For honorary volunteers
2. "Activos" - For active volunteers

Each sheet should have:
- First column: "Volunteer" (volunteer names)
- Subsequent columns: Course names

## Usage

1. Upload your Excel file using the file uploader
2. Select the volunteer type (Honorarios or Activos)
3. Choose the course from the dropdown menu
4. Enter the volunteer's name
5. Select the course completion date
6. Click "Update Course Information" to save the changes

The application will automatically update the Excel file with the new information. 
