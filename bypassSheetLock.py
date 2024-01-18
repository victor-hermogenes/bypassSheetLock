import PySimpleGUI as sg
import pandas as pd
import zipfile
import os
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
from shutil import rmtree

# Function to remove sheet protection by modifying the XML
def remove_sheet_protection(input_file_path, sheet_name):
    temp_dir = mkdtemp()
    # Unzip the .xlsx file
    with zipfile.ZipFile(input_file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Find the sheet file in the unzipped directory
    workbook_rel_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
    workbook_rel_tree = ET.parse(workbook_rel_path)
    sheet_file_name = None
    for element in workbook_rel_tree.getroot():
        if 'Target' in element.attrib:
            target = element.attrib['Target']
            if target.startswith('worksheets/') and sheet_name in target:
                sheet_file_name = target
                break
    
    if sheet_file_name is None:
        raise Exception(f"Sheet named '{sheet_name}' not found in the workbook.")
    
    # Remove the sheetProtection tag from the sheet XML
    sheet_path = os.path.join(temp_dir, 'xl', sheet_file_name)
    sheet_tree = ET.parse(sheet_path)
    sheet_root = sheet_tree.getroot()
    namespace = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    protection_tag = sheet_root.find(f'{namespace}sheetProtection')
    if protection_tag is not None:
        sheet_root.remove(protection_tag)
        sheet_tree.write(sheet_path)
    
    # Re-zip the contents back into a .xlsx file
    with zipfile.ZipFile(input_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                file_path_in_zip = os.path.relpath(file_path, temp_dir)
                zip_ref.write(file_path, file_path_in_zip)
    
    # Cleanup the temporary directory
    rmtree(temp_dir)

# Function to get sheet names using pandas
def get_sheet_names(excel_path):
    xl = pd.ExcelFile(excel_path)
    return xl.sheet_names

# GUI layout to select the Excel file
layout = [
    [sg.Text("Select Excel File"), sg.Input(), sg.FileBrowse(key="FILE", file_types=(("Excel Files", "*.xlsx"), ("Excel Workbook", "*.xlsb")))],
    [sg.Button("Load Sheets")]
]

window = sg.Window("Excel Sheet Unprotector", layout)

# Event loop for file selection
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Load Sheets":
        file_path = values["FILE"]
        if file_path:
            sheet_names = get_sheet_names(file_path)
            sheet_selection_layout = [
                [sg.Listbox(values=sheet_names, size=(20, 12), key="SHEET_NAME", enable_events=True)],
                [sg.Button("Unprotect Sheet")]
            ]
            window.close()
            window = sg.Window("Select a Sheet to Unprotect", sheet_selection_layout)
        else:
            sg.popup("Please select an Excel file.")
        break

# Event loop for sheet selection and unprotection
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Unprotect Sheet":
        selected_sheet_name = values["SHEET_NAME"][0]  # Assuming the user selects a sheet
        try:
            remove_sheet_protection(file_path, selected_sheet_name)
            sg.popup(f"Successfully unprotected '{selected_sheet_name}'!")
        except Exception as e:
            sg.popup(f"An error occurred: {e}")
        break

window.close()
