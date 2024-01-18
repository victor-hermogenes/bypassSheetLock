import PySimpleGUI as sg
import pandas as pd
import zipfile
import os
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
from shutil import rmtree

def unprotect_sheet(input_file_path, sheet_id):
    temp_dir = mkdtemp()
    with zipfile.ZipFile(input_file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # The sheet file name convention is 'sheet[id].xml' such as 'sheet1.xml'
    sheet_file_name = f'sheet{sheet_id}.xml'
    sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', sheet_file_name)
    
    # Parse the sheet XML and remove the sheetProtection element
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    namespace = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    protection_tag = root.find(f'{namespace}sheetProtection')
    if protection_tag is not None:
        root.remove(protection_tag)
        tree.write(sheet_path)
    
    # Re-zip the contents back into a .xlsx file
    with zipfile.ZipFile(input_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                file_path_in_zip = os.path.relpath(file_path, temp_dir)
                zip_ref.write(file_path, file_path_in_zip)
    
    rmtree(temp_dir)

def get_sheet_names(excel_path):
    xl = pd.ExcelFile(excel_path)
    return xl.sheet_names

# GUI layout to select the Excel file and list sheet names
layout = [
    [sg.Text("Select Excel File"), sg.Input(), sg.FileBrowse(key="FILE")],
    [sg.Button("Load Sheets")],
]

window = sg.Window("Excel Sheet Unprotector", layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Load Sheets":
        file_path = values["FILE"]
        if file_path:
            sheet_names = get_sheet_names(file_path)
            # Create numbered sheet list
            numbered_sheets = [f"{i+1}: {name}" for i, name in enumerate(sheet_names)]
            window.close()
            layout = [
                [sg.Text("Select sheet ID to unprotect:")],
                [sg.Listbox(values=numbered_sheets, size=(20, 12), key="SHEET_ID")],
                [sg.Button("Unprotect Sheet")],
            ]
            window = sg.Window("Select a Sheet ID to Unprotect", layout)
        else:
            sg.popup("Please select an Excel file.")
        break

# Event loop for sheet selection and unprotection
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Unprotect Sheet":
        selected_sheet = values["SHEET_ID"][0] if values["SHEET_ID"] else None
        if selected_sheet:
            # Extract the ID from the selected sheet
            sheet_id = int(selected_sheet.split(":")[0])
            try:
                unprotect_sheet(file_path, sheet_id)
                sg.popup(f"Sheet ID {sheet_id} has been unprotected.")
            except Exception as e:
                sg.popup(f"An error occurred: {e}")
        else:
            sg.popup("Please select a sheet ID.")
        break

window.close()

