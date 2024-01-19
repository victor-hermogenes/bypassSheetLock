import PySimpleGUI as sg
from removefunctions import remove_sheet_protection as rsp
from removefunctions import remove_workbook_protection as rwp

# Create interface
layout = [
    [sg.Text("Please, this program were created to unlock sheets that you forgot!")],
    [sg.Text("Any other ilegal uses will be on user's own responsability.")],
    [sg.Text("File Path: "), sg.Input(), sg.FileBrowse()],
    [sg.Button("Unprotect"), sg.Button("Close")],
    [sg.Text("Created by Victor G. Hermogenes")],
]

# Create window that will open to show interface
window = sg.Window("bypassSheetLock", layout)

# Creating events for buttons
while True:
    event, values = window.read()
    # Closes the window
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    # Run unprotect functions
    elif event == "Unprotect":
        file_path = values[0]
        rwp(file_path)
        rsp(file_path)
    else:
        break

window.close()

