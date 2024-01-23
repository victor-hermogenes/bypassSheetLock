import PySimpleGUI as sg
from removefunctions import remove_workbook_protection as rwp

# Callback function to update the progress bar and output box
def callback(progress_bar, output_box, message, progress):
    progress_bar.UpdateBar(progress)
    output_box.Update(f'{message}\n', append=True)

# Create interface
main_layout = [
    [sg.Text("Please, this program was created to unlock sheets that you forgot!")],
    [sg.Text("Any other illegal uses will be on user's own responsibility.")],
    [sg.Text("File Path: "), sg.Input(), sg.FileBrowse()],
    [sg.Button("Unprotect"), sg.Button("Close")],
    [sg.Text("Created by Victor G. Hermogenes")],
]

# Create window that will open to show interface
main_window = sg.Window("bypassSheetLock", main_layout)

# Event loop
try:
    while True:
        event, values = main_window.read()

        if event == sg.WIN_CLOSED or event == 'Close':
            break
        elif event == "Unprotect":
            file_path = values[0]

            layout_progress = [
                [sg.Text("Progress")],
                [sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')],
                [sg.Multiline(default_text='Starting...\n', size=(40, 10), key='output')],
            ]
            popup_window = sg.Window("Unprotecting...", layout_progress).finalize()
            progress_bar = popup_window['progressbar']
            output_box = popup_window['output']

            rwp(file_path, lambda msg, p=0: callback(progress_bar, output_box, msg, p))

            popup_window.close()

except Exception as e:
    sg.PopupError(f"Error: {e}")
finally:
    main_window.close()
