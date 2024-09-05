# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer

import tkinter
import tkinter.messagebox
import customtkinter
import tkinterDnD
import threading
from tkinter.filedialog import askopenfile

import win32com.client

from modules import fill_block as fb

customtkinter.set_ctk_parent_class(tkinterDnD.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("800x350")
app.title("a_acad v.0.001 (temp)")
# app.resizable(width=False, height=False)

# print(type(app), isinstance(app, tkinterDnD.Tk))

def button_callback():
    print("Button click", btn_filler._text)

def btn_exj_callback():
    print("Button click", btn_exj_chart._text)

"""Browse file dialog"""
content = "Choose tag-list..."
def open_file():
    global content
    file = askopenfile(mode="r")  # there is an option to choose only .xlsx- filetypes=[("Excel Files", "*.xlsx")]
    if file is not None:
        content = file.name
        print(content)
        file_name = customtkinter.CTkLabel(master=frame_1, text=f"{content}", wraplength=300)
        file_name.grid(row=2, column=0, pady=10, padx=0.5)
    return content


side_frame = customtkinter.CTkFrame(master=app, width=150, corner_radius=10)
side_frame.grid(column=0, rowspan=4, pady=20, padx=20, sticky="nsew")
side_frame.grid_rowconfigure(4, weight=1)
side_label = customtkinter.CTkLabel(master=side_frame, text="Modules", anchor="w",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
side_label.pack(pady=10, padx=10, anchor="w")


btn_filler = customtkinter.CTkButton(master=side_frame, command=button_callback, text="Block check")
btn_filler.pack(pady=2, padx=2)

btn_exj_chart = customtkinter.CTkButton(master=side_frame, command=btn_exj_callback, text="EXJ chart")
btn_exj_chart.pack(pady=4, padx=2)


frame_1 = customtkinter.CTkFrame(master=app, width=120, corner_radius=10)
frame_1.grid(row=0, column=1, rowspan=4, sticky="nsew", pady=20, padx=5)

label_1 = customtkinter.CTkLabel(master=frame_1, text="Choose support tag list file:", anchor="w",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_1.grid(pady=10, padx=10, sticky="nsew")

browse_button = customtkinter.CTkButton(master=frame_1, fg_color="transparent", border_width=2,
                                                     text="Browse file", text_color=("gray10", "#DCE4EE"),
                                                     command=open_file)
browse_button.grid(row=2, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
file_name = customtkinter.CTkLabel(master=frame_1, text=f"{content}")
file_name.grid(row=2, column=0, pady=10, padx=1)

progress_bar = customtkinter.CTkLabel(master=app, text=f"")
progress_bar.grid(row=5, column=1, pady=10, padx=0.5)
result_filling = customtkinter.CTkLabel(master=app, text=f"")
result_filling.grid(row=6, column=1, pady=10, padx=0.5)

def run_check():
    fb.check_sd_tags(content)
    result_checking = customtkinter.CTkLabel(master=frame_1, text=f"Checking complete!\n Result_1 \n Result_2")
    result_checking.grid(row=3, column=0, pady=10, padx=0.5)
    return

def run_fill():
    result_filling = customtkinter.CTkLabel(master=app, text=f"")
    result_filling.grid(row=6, column=1, pady=10, padx=0.5)
    fb.fill_support_tags(content)
    result_filling = customtkinter.CTkLabel(master=app, text=f"Filling complete!\n Result_1 \n Result_2")
    result_filling.grid(row=6, column=1, pady=10, padx=0.5)
    return

def fill_support_tags_test():
    acad = win32com.client.Dispatch("AutoCAD.Application")

    doc = acad.ActiveDocument  # Document object
    tag_list = fb.read_tags_pd(content)

    progress = 0
    for entity in acad.ActiveDocument.PaperSpace:
        name = entity.EntityName

        progress_percentage = f"{round(((progress / len(acad.ActiveDocument.PaperSpace)) * 100), 0)} %"
        progress_bar = customtkinter.CTkLabel(master=app, text=f"{progress_percentage}")
        progress_bar.grid(row=5, column=1, pady=10, padx=0.5)

        progress += 1
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                # print(entity.Name)
                # print(entity.Layer)
                # print(entity.ObjectID)
                support_tag = "n/a"
                for attrib in entity.GetAttributes():
                    # print("******")
                    # print("  {}: {}".format(attrib.TagString, attrib.TextString))
                    if "TAG" in attrib.TagString:
                        # print("!!!!!tag ", attrib.TextString)
                        support_tag = attrib.TextString.strip()
                    # update text
                    # if 'DRAWING' in attrib.TagString and "LATER" in attrib.TextString:
                    if 'DRAWING' in attrib.TagString:
                        # print(f"  --- {attrib.TextString}")
                        try:
                            attrib.TextString = tag_list[support_tag]
                            attrib.Update()
                        except Exception as e:
                            print(e)
                            pass
    result_filling = customtkinter.CTkLabel(master=app, text=f"Filling SD-TAGs complete!")
    result_filling.grid(row=6, column=1, pady=10, padx=0.5)
    return


def start_fill():
    result_filling.destroy()
    progress_bar.destroy()

    btn_fill_tags.configure(state=tkinter.DISABLED)
    thread = threading.Thread(target=fill_support_tags_test)
    print(threading.main_thread().name)
    print(thread.name)
    thread.start()
    check_thread(thread)
    return

def check_thread(thread):
    if thread.is_alive():
        app.after(100, lambda: check_thread(thread))
    else:
        print(thread.is_alive())
        btn_fill_tags.configure(state=tkinter.NORMAL)


btn_check_tags = customtkinter.CTkButton(master=app, command=run_check, text="Check tags")
btn_check_tags.grid(row=4, column=1, pady=10, padx=0.1, sticky="w")

btn_fill_tags = customtkinter.CTkButton(master=app, command=start_fill, text="Fill tags")
btn_fill_tags.grid(row=5, column=1, pady=10, padx=0.1, sticky="w")






# progressbar_1 = customtkinter.CTkProgressBar(master=frame_1)
# progressbar_1.pack(pady=10, padx=10)
#
# button_1 = customtkinter.CTkButton(master=frame_1, command=button_callback)
# button_1.pack(pady=10, padx=10)
#
# slider_1 = customtkinter.CTkSlider(master=frame_1, command=slider_callback, from_=0, to=1)
# slider_1.pack(pady=10, padx=10)
# slider_1.set(0.5)
#
# entry_1 = customtkinter.CTkEntry(master=frame_1, placeholder_text="CTkEntry")
# entry_1.pack(pady=10, padx=10)

# optionmenu_1 = customtkinter.CTkOptionMenu(frame_1, values=["Option 1", "Option 2", "Option 42 long long long..."])
# optionmenu_1.pack(pady=10, padx=10)
# optionmenu_1.set("CTkOptionMenu")
#
# combobox_1 = customtkinter.CTkComboBox(frame_1, values=["Option 1", "Option 2", "Option 42 long long long..."])
# combobox_1.pack(pady=10, padx=10)
# combobox_1.set("CTkComboBox")
#
# checkbox_1 = customtkinter.CTkCheckBox(master=frame_1)
# checkbox_1.pack(pady=10, padx=10)
#
# radiobutton_var = customtkinter.IntVar(value=1)
#
# radiobutton_1 = customtkinter.CTkRadioButton(master=frame_1, variable=radiobutton_var, value=1)
# radiobutton_1.pack(pady=10, padx=10)
#
# radiobutton_2 = customtkinter.CTkRadioButton(master=frame_1, variable=radiobutton_var, value=2)
# radiobutton_2.pack(pady=10, padx=10)
#
# switch_1 = customtkinter.CTkSwitch(master=frame_1)
# switch_1.pack(pady=10, padx=10)
#
# text_1 = customtkinter.CTkTextbox(master=frame_1, width=200, height=70)
# text_1.pack(pady=10, padx=10)
# text_1.insert("0.0", "CTkTextbox\n\n\n\n")
#
# segmented_button_1 = customtkinter.CTkSegmentedButton(master=frame_1, values=["CTkSegmentedButton", "Value 2"])
# segmented_button_1.pack(pady=10, padx=10)

# tabview_1 = customtkinter.CTkTabview(master=frame_1, width=300, )
# tabview_1.pack(pady=10, padx=10)
# tabview_1.add("CTkTabview")
# tabview_1.add("Tab 2")

app.mainloop()