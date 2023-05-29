import PySimpleGUI as sg  # Part 1 - The import
from exportWord import exportword
# Define the window's contents
layout = [[sg.Text("What's your new filename?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("Which year?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("Which month?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("Which day?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("What's your old filepath?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("What's your combine filename?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text("What's your budget filename?")],  # Part 2 - The Layout
          [sg.Input()],
[sg.Text(size=(40,1), key='-OUTPUT-')],
          [sg.Button('Ok'), sg.Button('Quit')]]

# Create the window
window = sg.Window('File Creat', layout)  # Part 3 - Window Defintion

while True:
    event, values = window.read()
    # See if user wants to quit or window was closed
    if event in [sg.WINDOW_CLOSED, 'Quit']:
        break
    # Output a message to the window
    args = list(values.values())
    num = sum(1 for arg in args if len(arg))
    if num!= 7:
        window['-OUTPUT-'].update("Error: args need all! Retry!")
    exportword(*args)
    window['-OUTPUT-'].update("File created!")
window.close()