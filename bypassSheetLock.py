import PySimpleGUI as sg
import os

def unprotect_excel(file_path):
    extension = os.path.splitext(file_path)[1]
    if extension in ['.xlsx', '.xls']:
        # Add logic to unprotect Excel files
        pass
    elif extension == '.csv':
        # Add logic if needed for CSV files
        pass
    else:
        sg.popup("Unsupported File Format", f"The file {file_path} is not a supported format.")

layout = [
    [sg.Text("Select Excel File")],
    [sg.Input(), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")))],
    [sg.Button("Unprotect"), sg.Exit()]
]

window = sg.Window("Excel Unprotecter", layout)

while True:
    event, values = window.read()
    if event in (None, "Exit"):
        break
    if event == "Unprotect":
        file_path = values[0]
        if file_path:
            unprotect_excel(file_path)
            sg.popup("Process Completed", "The file has been processed.")
        else:
            sg.popup("Error", "Please select a file.")

window.close()
