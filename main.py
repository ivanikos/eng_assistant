# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer

import os
import eel

import fill_block as fb

@eel.expose
def read_path():

        return




if __name__ == '__main__':
    browser_path = (r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\cad_helper"
                    r"\chrome-win\chrome.exe")

    eel.init('front')
    eel.start('index.html', mode="edge", size=(900, 780), 
              cmdline_args="start msedge --new-window --app=https://localhost:8000")


