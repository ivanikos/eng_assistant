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
from tkinter.filedialog import askopenfile, askdirectory

import win32com.client

from modules.fill_block import check_sd_tags
from modules.logics import write_log
from modules.sm_logic import get_pcat_list, change_spec_paths

# Functions section-------------------------------------------------------------

content = "Select .pspx file ..."
pcat_paths ={}
original_pcat_paths_list = {}
not_ok = []

def check_compatibility(catalog_list, new_directory_path):
    for catalog_name in pcat_paths:
        if catalog_name in catalog_list:
            pcat_paths[catalog_name][0] = f"{new_directory_path}\\{catalog_name}"
            pcat_paths[catalog_name][2] = "new"

    pcat_textbox.configure(state="normal")  # Enable editing to append paths
    pcat_textbox.delete("1.0", "end-1c")
    pcat_textbox.configure(state="disabled")  # Set back to read-only

    check_status = 1
    if not pcat_paths:
        check_status = 0
    for i in pcat_paths:
        print(i, pcat_paths[i])

        if pcat_paths[i][2] == "new":
            pcat_textbox.configure(state="normal")  # Enable editing to append paths
            pcat_textbox.insert("end", f"{pcat_paths[i][0].replace("/", "\\")} - NEW PATH OK!\n\n")  # Add path to the text field
            pcat_textbox.configure(state="disabled")  # Set back to read-only

        if pcat_paths[i][2] == "old":
            check_status = 0
            pcat_textbox.configure(state="normal")  # Enable editing to append paths
            pcat_textbox.insert("end", f"Catalog {i} NOT FOUND, select correct path to catalog manually\n\n")  # Add path to the text field
            pcat_textbox.configure(state="disabled")  # Set back to read-only

    return check_status
def manual_catalog_path(button):
    # button_text = button.cget("text")
    print(button)

def pop_paths_error():
    global pcat_paths
    # Create a new top-level window for help
    path_error_window = customtkinter.CTkToplevel(app)
    path_error_window.title("Migration Status")
    path_error_window.after(250, lambda: path_error_window.iconbitmap(r'./icons/icon.ico'))
    path_error_window.geometry("1000x200")  # Set the desired window size
    path_error_window.grid_columnconfigure(1, weight=1)
    # path_error_window.resizable(False, False)
    path_error_window.attributes('-topmost', True)

    help_message = (
        "Manual selecting path:"
    )

    frame_help = customtkinter.CTkFrame(master=path_error_window, corner_radius=10)
    frame_help.grid(row=0, column=0, rowspan=4, columnspan=2, sticky="nsew", pady=10, padx=10)

    label_top = customtkinter.CTkLabel(master=frame_help, text="P3D Spec Migration tool",
                                       font=customtkinter.CTkFont(size=20, weight="bold"))
    label_top.grid(row=0, columnspan=2, pady=10, padx=10, sticky="nsew")

    label_mid = customtkinter.CTkLabel(master=frame_help, text=f"{help_message}",
                                       font=customtkinter.CTkFont(size=14))
    label_mid.grid(row=1, columnspan=2, pady=10, padx=10, sticky="nsew")

    row_number = 5
    for i in pcat_paths:
        if pcat_paths[i][2] == "old":
            wrong_path = customtkinter.CTkButton(master=path_error_window, fg_color="transparent",
                                                 border_width=2, text=f"{i}", command=lambda path=i: manual_catalog_path(path))
            wrong_path.grid(row=f"{row_number}", columnspan=2, pady=5, padx=5, sticky="w")
            row_number += 1

    # Create action buttons
    close_button = customtkinter.CTkButton(master=path_error_window, text="Done", command=path_error_window.destroy)
    close_button.grid(row=f"{row_number}", column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

def open_pcat_folder():
    global pcat_paths
    catalog_list = []
    folder_path = askdirectory()
    migrate_button.configure(state=tkinter.DISABLED)

    try:
        pcat_paths_list = get_pcat_list(content)[0]
        for i in range(0, (len(pcat_paths_list))):
            pcat_paths[pcat_paths_list[i].split("\\")[-1]] = [pcat_paths_list[i], "", "old"]

        if folder_path:
            print(os.listdir(folder_path))
            print(f"New .pcat files directory: {folder_path}")
            catalog_list = [i for i in os.listdir(folder_path) if ".pcat" in i]
            for i in catalog_list:
                print(i)
            if check_compatibility(catalog_list, folder_path) == 1:
                migrate_button.configure(state=tkinter.NORMAL)
                pcat_textbox.delete("1.0", "end-1c")
            else:
                manual_browse_button_pcat.configure(state=tkinter.NORMAL)
    except:
        print("error")
        pcat_textbox.configure(state="normal")  # Enable editing to append paths
        pcat_textbox.delete("1.0", "end-1c")
        pcat_textbox.insert("1.0", "ERROR:  No such file or directory.")
        pcat_textbox.configure(state="disabled")





def open_file():
    """Browse file dialog"""
    global content
    global pcat_paths
    global file_name
    global label_pcat
    global pcat_textbox
    pcat_paths = {}
    content = ""

    migrate_button.configure(state=tkinter.DISABLED)

    file = askopenfile(mode="r", filetypes=[("Spec Files", "*.pspx")])  # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    if file is not None:
        content = file.name
        pcat_paths_list = get_pcat_list(content)[0]

        file_name.destroy()
        file_name = customtkinter.CTkLabel(master=frame_spec_path, text=f"{content}", wraplength=800)
        file_name.grid(row=2, column=0, pady=10, padx=10, sticky="w")

        pcat_textbox.destroy()
        pcat_textbox = customtkinter.CTkTextbox(master=frame_pcat, wrap="none", height=(35 * len(pcat_paths_list)),
                                                width=600)
        print(35 * len(pcat_paths_list), len(pcat_paths_list))
        pcat_textbox.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")
        app.update_idletasks()
        app.geometry("1050x300")

        try:
            pcat_paths_list = get_pcat_list(content)[0]
            app.geometry(f"1050x{400 + (25 * len(pcat_paths_list))}")
            print(f"1050x{400 + (25 * len(pcat_paths_list))}")

            for i in range(0, (len(pcat_paths_list))):
                pcat_textbox.configure(state="normal")  # Enable editing to append paths
                pcat_textbox.insert("end", f"{pcat_paths_list[i]}\n\n")  # Add path to the text field
                pcat_textbox.configure(state="disabled")  # Set back to read-only

                pcat_paths[pcat_paths_list[i].split("\\")[-1]] = [pcat_paths_list[i], "", "old"]
                original_pcat_paths_list[pcat_paths_list[i].split("\\")[-1]] = [pcat_paths_list[i], ""]

                app.update_idletasks()

        except Exception as e:
            return f"ERROR: {e}"
    for i in pcat_paths:
        print(i, pcat_paths[i])
    return content


# Functions section end -------------------------------------------------------------
def migrate():
    try:
        change_spec_paths(content, pcat_paths)
        pop_window_menu_action()
        migrate_button.configure(state=tkinter.DISABLED)
    except Exception as e:
        pass

def pop_window_menu_action():
    # Create a new top-level window for help
    help_window = customtkinter.CTkToplevel(app)
    help_window.title("Migration Status")
    help_window.after(250, lambda: help_window.iconbitmap(r'./icons/icon.ico'))
    help_window.geometry("450x200")  # Set the desired window size
    help_window.grid_columnconfigure(1, weight=1)
    help_window.resizable(False, False)
    # help_window.focus_force()
    # Help message content
    help_window.attributes('-topmost', True)

    help_message = (
        "Migration successful!"
    )

    frame_help = customtkinter.CTkFrame(master=help_window, corner_radius=10)
    frame_help.grid(row=0, column=0, rowspan=4, columnspan=2, sticky="nsew", pady=10, padx=10)

    label_top = customtkinter.CTkLabel(master=frame_help, text="P3D Spec Migration tool",
                                       font=customtkinter.CTkFont(size=20, weight="bold"))
    label_top.grid(row=0, columnspan=2, pady=10, padx=10, sticky="nsew")

    label_mid = customtkinter.CTkLabel(master=frame_help, text=f"{help_message}",
                                       font=customtkinter.CTkFont(size=14))
    label_mid.grid(row=1, columnspan=2, pady=10, padx=10, sticky="nsew")

    # Create action buttons
    close_button = customtkinter.CTkButton(help_window, text="OK", command=help_window.destroy)
    close_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")


# APP initialisation ------------------------------------------
customtkinter.set_ctk_parent_class(tkinterdnd2.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("1050x350")
app.eval('tk::PlaceWindow . center')
app.update_idletasks()
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
                                                        text="Browse new catalogs folder", text_color=("gray10", "#DCE4EE"),
                                                        command=open_pcat_folder)
browse_button_pcat.grid(row=0, column=1, padx=30, pady=10, sticky="e")

manual_browse_button_pcat = customtkinter.CTkButton(master=frame_pcat, fg_color="transparent", border_width=2,
                                                        text="MANUALLY Browse new catalog paths", text_color=("gray10", "#DCE4EE"),
                                                        command=pop_paths_error)
manual_browse_button_pcat.grid(row=1, column=1, padx=30, pady=10, sticky="e")
manual_browse_button_pcat.configure(state=tkinter.DISABLED)



pcat_textbox = customtkinter.CTkTextbox(master=frame_pcat, wrap="none", height=50, width=600)
pcat_textbox.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")
pcat_textbox.configure(state="disabled")  # Make it read-only initially

label_placeholder = customtkinter.CTkLabel(master=frame_pcat, text="",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_placeholder.grid(row=2, column=0, padx=2, pady=2, sticky="w")

migrate_button = customtkinter.CTkButton(master=frame_spec_path, command=migrate, text="Migrate!")
migrate_button.grid(row=3, column=1, padx=30, pady=10, sticky="nsew")
migrate_button.configure(state=tkinter.DISABLED)

# compatibility_button = customtkinter.CTkButton(master=frame_spec_path, text="Check compatibility")
# compatibility_button.grid(row=3, column=0, padx=10, pady=10, sticky="e")
# compatibility_button.configure(state=tkinter.DISABLED)

app.mainloop()







