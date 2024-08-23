import PySimpleGUI as sg


layout = [[sg.Text(text='a_acad',
                    font=('Arial Bold', 20),
                    size=20,
                    expand_x=True,
                    justification='left')],

            [sg.Text(text='Choose Pipe Support tag list:', k="-T-",
                    font=('Arial', 16),
                    size=20,
                    expand_x=True,
                    justification='left'), ],

            [sg.Combo(sorted(sg.user_settings_get_entry('-filenames-', [])), k="-A-",
                    default_value=sg.user_settings_get_entry('-last filename-', ''),
                      size=(50, 1), key='-FILENAME-'), sg.FileBrowse(), sg.B('Check Tags'), sg.B('Fill Tags')],



          [sg.Button('Ok'), sg.Button('Cancel')]

          ]

layout_1 = [[sg.Text(text='a_acad',
                    font=('Arial Bold', 20),
                    size=20,
                    expand_x=True,
                    justification='left')],
            [sg.Text(text='new layout',
                    font=('Arial Bold', 20),
                    size=20,
                    expand_x=True,
                    justification='left')],

          [sg.Button('Ok'), sg.Button('Cancel')]

          ]