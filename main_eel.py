# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer

# import os
# import eel
#
# import fill_block as fb
#
# @eel.expose
# def read_path():
#
#         return
#
#
#
#
# if __name__ == '__main__':
#     browser_path = (r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\cad_helper"
#                     r"\chrome-win\chrome.exe")
#
#     eel.init('front')
#     eel.start('index.html', mode="edge", size=(900, 780),
#               cmdline_args="start chrome --new-window --app=https://localhost:8000")
#
#

import tkinter
import tkinter.messagebox
from tkinter.filedialog import askopenfile
import customtkinter
from openpyxl.styles.builtins import comma
from win32pdhutil import browse

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


"""Browse file dialog"""
content = "Choose tag-list..."
def open_file():
    global content
    file = askopenfile(mode="r")  # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    if file is not None:
        content = file.name
        print(content)
        app.chosen_file_name._text = f"{content}"
    return content


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("a_acad v.0.001 (temp)")
        self.geometry(f"{800}x{300}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        # self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Modules",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_1_event, text="Blocks checking")
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_2_event, text="EXJ chart")
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)


        # create main entry and button
        self.browse_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text="Browse file", text_color=("gray10", "#DCE4EE"),
                                                     command=self.chosing_file)
        self.browse_button.grid(row=2, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.choose_file = customtkinter.CTkLabel(master=self, text="Choose Support tags list file:",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.choose_file.grid(row=1, column=1, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.chosen_file_name = customtkinter.CTkLabel(master=self, font=customtkinter.CTkFont(size=12),
                                                       text=f"{content}", )
        self.chosen_file_name.grid(row=2, column=1, padx=(20, 20), pady=(20, 20), sticky="nsew")



    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_1_event(self):
        print(f"sidebar_button {self.sidebar_button_1._text} click")

    def sidebar_button_2_event(self):
        print(f"sidebar_button {self.sidebar_button_2._text} click")





if __name__ == "__main__":
    app = App()
    app.mainloop()