import PySimpleGUI as sg
from pywin.scintilla.view import event_commands

import fill_block as fb
import layouts


layout = layouts.layout

window = sg.Window('a_acad v.0.01 (temp)', layout, size=(None, None),
                    enable_close_attempted_event=True,
                   location=sg.user_settings_get_entry('-location-', (None, None)))

while True:
    event, values = window.read()
    print(f" event - {event}, values - {values}")

    # need to add pop-up reminder if there is no tag-list file

    if event == "Fill Tags":
        if values["-FILENAME-"]:
            fb.fill_support_tags(values["-FILENAME-"])

    if event == "Check Tags":
        if values["-FILENAME-"]:
            fb.check_sd_tags(values["-FILENAME-"])

    if event == "Ok":
        sg.user_settings_set_entry('-location-', window.current_location())
        window.close()
        window = sg.Window('AutoCAD Helper v.0.0 (temp)', layouts.layout_1, size=(715, 250),
                          enable_close_attempted_event=True,
                          location=sg.user_settings_get_entry('-location-', (None, None)))

    if event in (None, "Exit", "Cancel", "-WINDOW CLOSE ATTEMPTED-"):
        sg.user_settings_set_entry('-location-', window.current_location())
        window.close()
        break
