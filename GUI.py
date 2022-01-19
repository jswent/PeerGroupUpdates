import PySimpleGUI as sg

retur = ""

layout = [[sg.Text('Pull Peer Group Information:')],
          [sg.In(key='-CAL-', enable_events=False, visible=False),
           sg.CalendarButton('Select Date to Pull From', target='-CAL-', pad=None, font=('MS Sans Serif', 10, 'bold'),
                             button_color=('red', 'white'), key='_CALENDAR_', format=('%d %B, %Y'))],
          [sg.In(), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
          [sg.Exit()]]

window = sg.Window('Calendar', layout)

while True:  # Event Loop
    event, values = window.read()
    retur = (event, values)
    if event in (None, 'Exit'):
        break
window.close()


def returnParam():
    return retur
