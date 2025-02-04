# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer
import os
import tkinter
import tkinter.messagebox
from modulefinder import packagePathMap

import customtkinter
import tkinterdnd2
import threading
from tkinter.filedialog import askopenfile, asksaveasfilename

from modules.sm_logic import get_pcat_list, change_spec_paths

# Functions section-------------------------------------------------------------

content = "Select .pspx file ..."
pcat_paths ={}
original_pcat_paths_list = {}
not_ok = []
def open_new_pcat():
    global pcat_paths
    global original_pcat_paths_list
    global not_ok
    pcat_file = askopenfile(mode="r", filetypes=[("Catalog Files", "*.*cat")])  # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    print(pcat_file.name)


    if pcat_file is not None:
        try:
            row_number = pcat_paths[str(pcat_file.name).split("/")[-1]][1]

            pcat_paths[str(pcat_file.name).split("/")[-1]][0] = pcat_file.name
            pcat_paths[str(pcat_file.name).split("/")[-1]][2] = "ok"

            file_name_pcat = customtkinter.CTkLabel(master=frame_pcat, text=pcat_file.name.replace("/", "\\"), anchor="center")
            file_name_pcat.grid(row=f"{row_number}", column=0,
                                pady=10, padx=10, sticky="w")

            status_label = customtkinter.CTkLabel(master=frame_pcat, text="OK!")
            status_label.grid(row=row_number, column=1, padx=30, pady=10, sticky="nsew")

        except Exception as e:
            print("NO related catalog in that spec")
            print(e)
            print(pcat_file)

    not_ok = []
    for i in pcat_paths:
        if pcat_paths[i][2] == "old":
            not_ok.append(pcat_paths[i][0])
    if not not_ok:
        migrate_button.configure(state=tkinter.NORMAL)


def open_file():
    """Browse file dialog"""
    global content
    global pcat_paths
    global file_name
    global label_pcat
    pcat_paths = {}

    file_name.destroy()
    file_name = customtkinter.CTkLabel(master=frame_spec_path, text=f"{content}", wraplength=800)
    file_name.grid(row=2, column=0, pady=10, padx=10, sticky="w")
    app.update_idletasks()

    file = askopenfile(mode="r", filetypes=[("Spec Files", "*.pspx")])  # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    if file is not None:
        content = file.name
        file_name = customtkinter.CTkLabel(master=frame_spec_path, text=f"{content}", wraplength=800)
        file_name.grid(row=2, column=0, pady=10, padx=10, sticky="w")

        app.geometry("1050x300")
        frame_pcat = customtkinter.CTkFrame(master=app, corner_radius=10)
        frame_pcat.grid(row=4, column=1, rowspan=3, sticky="nsew", padx=10, pady=10)
        frame_pcat.columnconfigure(0, weight=2)
        label_pcat = customtkinter.CTkLabel(master=frame_pcat, text="Related catalogs list:",
                                            font=customtkinter.CTkFont(size=20, weight="bold"))
        label_pcat.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        browse_button_pcat = customtkinter.CTkButton(master=frame_pcat, fg_color="transparent", border_width=2,
                                                     text="Browse new file", text_color=("gray10", "#DCE4EE"),
                                                     command=open_new_pcat)
        browse_button_pcat.grid(row=0, column=1, padx=30, pady=10, sticky="e")

        try:
            pcat_paths_list = get_pcat_list(content)[0]
            app.geometry(f"1050x{400 + (25 * len(pcat_paths_list))}")
            print(f"1050x{400 + (25 * len(pcat_paths_list))}")
            pcat_path_row = 1
            for i in range(0, (len(pcat_paths_list))):
                file_name_pcat = customtkinter.CTkLabel(master=frame_pcat, text=f"{pcat_paths_list[i]}", anchor="center")
                file_name_pcat.grid(row=f"{pcat_path_row}", column=0, pady=10, padx=10, sticky="w")

                pcat_paths[pcat_paths_list[i].split("\\")[-1]] = [pcat_paths_list[i], pcat_path_row, "old"]
                original_pcat_paths_list[pcat_paths_list[i].split("\\")[-1]] = [pcat_paths_list[i], pcat_path_row]

                app.update_idletasks()

                pcat_path_row += 1
        except Exception as e:
            return f"ERROR: {e}"
    for i in pcat_paths:
        print(i, pcat_paths[i])
    return content


# Functions section end -------------------------------------------------------------
def migrate():
    change_spec_paths(content, pcat_paths)


# APP initialisation ------------------------------------------
customtkinter.set_ctk_parent_class(tkinterdnd2.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("1050x300")
app.eval('tk::PlaceWindow . center')
# app.maxsize(1000, 300)
# app.minsize(1000, 300)
app.title("P3D Spec Migration tool")
app.iconbitmap(r'./icons/icon.ico')
app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(4, weight=1)
# APP initialisation end ------------------------------------------


# APP UI

frame_spec_path = customtkinter.CTkFrame(master=app, corner_radius=10)
frame_spec_path.grid(row=0, column=1, rowspan=4, sticky="nsew", pady=10, padx=10)
frame_spec_path.columnconfigure(0, weight=2)

user_name = os.getlogin()
user_name_ui = customtkinter.CTkLabel(master=frame_spec_path, text=f"User: {user_name}",
                                      font=customtkinter.CTkFont(size=10, weight="bold"))
user_name_ui.grid(row=0, column=1, padx=30, pady=5, sticky="e")

label_spec_path = customtkinter.CTkLabel(master=frame_spec_path, text="Open Plant 3D spec .pspx file:",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_spec_path.grid(row=0, column=0, pady=10, padx=10, sticky="w")

file_name = customtkinter.CTkLabel(master=frame_spec_path, text=f"{content}", anchor="center")
file_name.grid(row=2, column=0, pady=10, padx=10, sticky="w")
browse_button = customtkinter.CTkButton(master=frame_spec_path, fg_color="transparent", border_width=2,
                                                     text="Browse file", text_color=("gray10", "#DCE4EE"),
                                                     command=open_file)
browse_button.grid(row=2, column=1, padx=30, pady=10, sticky="nsew")


frame_pcat = customtkinter.CTkFrame(master=app, corner_radius=10)
frame_pcat.grid(row=4, column=1, rowspan=3, sticky="nsew", padx=10, pady=10)

frame_pcat.columnconfigure(0, weight=2)
label_pcat = customtkinter.CTkLabel(master=frame_pcat, text="Related catalogs list:",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_pcat.grid(row=0, column=0, padx=10, pady=10, sticky="w")
browse_button_pcat = customtkinter.CTkButton(master=frame_pcat, fg_color="transparent", border_width=2,
                                                        text="Browse new file", text_color=("gray10", "#DCE4EE"),
                                                        command=open_new_pcat)
browse_button_pcat.grid(row=0, column=1, padx=30, pady=10, sticky="e")


label_placeholder = customtkinter.CTkLabel(master=frame_pcat, text="",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_placeholder.grid(row=1, column=0, padx=10, pady=10, sticky="w")

migrate_button = customtkinter.CTkButton(master=frame_spec_path, command=migrate, text="Migrate!")
migrate_button.grid(row=3, column=1, padx=30, pady=10, sticky="nsew")
migrate_button.configure(state=tkinter.DISABLED)

app.mainloop()

