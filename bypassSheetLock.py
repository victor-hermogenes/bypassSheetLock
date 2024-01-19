import PySimpleGUI as sg
from removefunctions import remove_sheet_protection as rsp
from removefunctions import remove_workbook_protection as rwp


def upb(progress_bar, output_box, progress, message):
    """Update the progress bar and output box."""
    progress_bar.UpdateBar(progress)
    output_box.Update(f'{message}\n', append=True)


def layout_progress():
    layout_progress = [
        [sg.Text("Progress")],
        [sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')],
        [sg.Multiline(default_text='Events will appear here...', size=(40, 10), key='output')],
    ]
    
    return layout_progress

# Create interface
main_layout = [
    [sg.Text("Please, this program were created to unlock sheets that you forgot!")],
    [sg.Text("Any other ilegal uses will be on user's own responsability.")],
    [sg.Text("File Path: "), sg.Input(), sg.FileBrowse()],
    [sg.Button("Unprotect"), sg.Button("Close")],
    [sg.Text("Created by Victor G. Hermogenes")],
]



# Create window that will open to show interface
main_window = sg.Window("bypassSheetLock", main_layout)

# Creating events for buttons
try:
    while True:
        event, values = main_window.read()
        # Closes the window
        if event == sg.WIN_CLOSED or event == 'Close':
            break
        # Run unprotect functions
        elif event == "Unprotect":
            file_path = values[0]

            # Define a lambda function for the callback
            def callback(message, progress=0):
                upb(progress_bar, output_box, progress, message)
                popup_window.refresh()  # Refresh the window to update the GUI

            # Open popup window
            popup_window = sg.Window("Unprotection...", layout_progress()).finalize()
            progress_bar = popup_window['progressbar']
            output_box = popup_window['output']

            # Call functions with callback
            rwp(file_path, callback)
            rsp(file_path, callback)
        else:
            break

    main_window.close()
except Exception as e:
    print(f"Error: {e}")