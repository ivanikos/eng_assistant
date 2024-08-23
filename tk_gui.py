import tkinter as tk
from tabnanny import check
from tkinter.constants import CENTER
from tkinter.ttk import Label, Entry
from tkinter.filedialog import askopenfile

import fill_block as fb


window = tk.Tk()
window.title('a_acad v.0.001 (temp)')
window.geometry("800x300")
window.minsize(800, 300)
window.maxsize(800, 300)

""" Header menu"""
menu = tk.Menu(window)
window.config(menu=menu)

file_menu = tk.Menu(menu)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Option_1")
file_menu.add_command(label="Export report to Excel")
file_menu.add_separator()
file_menu.add_command(label="Exit", command=window.quit)

help_menu = tk.Menu(menu)
menu.add_cascade(label="Help", menu=help_menu)
help_menu.add_command(label="About")


# Label(window, text="text 1").grid(row=0)
# Label(window, text="text 2").grid(row=1)
# e_1.grid(row=0, column=1)
# e_2.grid(row=1, column=1)

Label(window, text="Choose Excel file with support tags list:",
      font="Arial 12").place(relx=0.03, rely=0.09)


"""Browse file dialog"""
def open_file():
    file = askopenfile(mode="r") # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    if file is not None:
        content = file.name
        Label(window, text=f"Your file \n {content}").place(relx=0.1, rely=0.25)

        print(content)
    return
browse_file_button= tk.Button(window, text="Browse", command=open_file)
browse_file_button.place(relx=0.03, rely=0.25)


"""Interface buttons"""
ok_button = tk.Button(window, text='Ok', width=15, command=window.quit)
cancel_button = tk.Button(window, text='Cancel', width=15, command=window.quit)
ok_button.place(relx=0.6, rely=0.85)
cancel_button.place(relx=0.8, rely=0.85)


"""Action button handler"""
def check_sd_tags(event):
    print(event)
    print("Button was clicked")

def fill_sd_tags(event):
    print(event)
    print("Button was clicked")

check_sd_tags_button = tk.Button(text="Check SD-Tags")
check_sd_tags_button.bind("<Button-1>", check_sd_tags)
check_sd_tags_button.place(relx=0.1, rely=0.4)

fill_sd_tags_button = tk.Button(text="fill SD-Tags")
fill_sd_tags_button.bind("<Button-1>", fill_sd_tags)
fill_sd_tags_button.place(relx=0.3, rely=0.4)






window.mainloop()