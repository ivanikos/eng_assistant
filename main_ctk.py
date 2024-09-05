# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer

import tkinter
import tkinter.messagebox
import customtkinter
import tkinterdnd2
import threading
from tkinter.filedialog import askopenfile

import win32com.client
import pandas as pd

from modules import fill_block as fb

customtkinter.set_ctk_parent_class(tkinterdnd2.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("800x400")
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
"""Read the tag-list .xlsm/xlsx etc."""
def read_tags_pd(path_tag_list_pd):
    tag_list = {}
    df = pd.read_excel(path_tag_list_pd, usecols=[2, 5, 11], sheet_name="Support Tags", header=7)

    for row in df.values:
        if row[0] != "nan" or "n":
            load_tag = row[0]
            project_number = str(row[1])[2: 10]
            sd_tag = str(row[2]).replace(f"-{project_number}", "").replace("nan", "")

            tag_list[load_tag] = sd_tag
            # print(project_number, load_tag, sd_tag)
        else:
            break
    return tag_list


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

progress_bar = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
progress_bar.grid(row=5, column=1, pady=0.2, padx=0.5)
result_filling = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
result_filling.grid(row=6, column=1, pady=0.2, padx=0.5)

progress_bar_checking = customtkinter.CTkLabel(master=app, text=f"")
progress_bar_checking.grid(row=5, column=1, pady=0.2, padx=0.5)
result_checking = customtkinter.CTkLabel(master=app, text=f"")
result_checking.grid(row=6, column=1, pady=0.2, padx=0.5)

def check_load_sd_tags():
    correct_tags = []
    wrong_load_tags = []
    wrong_sd_tags = []
    possible_to_fill = []
    waiting_load_tag = []
    waiting_sd_tag = []

    progress_bar_checking = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    progress_bar_checking.grid(row=5, column=1, pady=0.2, padx=0.5)
    result_checking = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    result_checking.grid(row=5, column=1, pady=2, padx=0.5)

    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument  # Document object
    tag_list = read_tags_pd(content)

    progress = 0
    for entity in acad.ActiveDocument.PaperSpace:
        name = entity.EntityName

        progress_percentage = f"{round(((progress / len(acad.ActiveDocument.PaperSpace)) * 100), 0)} %"
        progress_bar_checking = customtkinter.CTkLabel(master=app, text=f"{progress_percentage}")
        progress_bar_checking.grid(row=5, column=1, pady=5, padx=0.5)

        progress += 1
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                # print(entity.Name)
                # print(entity.Layer)
                # print(entity.ObjectID)
                load_tag = "n/a"
                for attrib in entity.GetAttributes():
                    if "DWGNO" in attrib.TagString:
                        print(attrib.TextString.strip())

                    if "TAG" in attrib.TagString:
                        load_tag = attrib.TextString.strip()


                    if "DRAWING" in attrib.TagString and attrib.TextString:
                        # print(load_tag, attrib.TextString)

                        sd_tag = str(attrib.TextString).strip()
                        try:
                            if str(load_tag).capitalize() == "LATER":
                                waiting_load_tag.append(f"{load_tag} - {sd_tag}")
                            elif load_tag in tag_list.keys() and tag_list[load_tag]:
                                if tag_list[load_tag] and sd_tag == "LATER":
                                    possible_to_fill.append(f"{load_tag} - {sd_tag} - "
                                                                f"{tag_list[load_tag]}")
                                elif tag_list[load_tag] and sd_tag != "LATER":
                                    if tag_list[load_tag].strip() == sd_tag.strip():
                                        correct_tags.append(f"{load_tag} - {sd_tag} - correct")
                                    else:
                                        wrong_sd_tags.append(f"{load_tag} - {sd_tag} - INcorrect")
                                else:
                                    print(f"ELSE - {load_tag} - {sd_tag}")
                            else:
                                if not tag_list[load_tag]:
                                    waiting_sd_tag.append(f"{load_tag} - {sd_tag} - wait sd-tag")
                                wrong_load_tags.append(f"{load_tag} - {sd_tag} - need to double check")
                        except Exception as e:
                            print(e)
                            # print(f"{attrib.TagString} -- {attrib.TextString} -- WRONG TAG")

    result_checking = customtkinter.CTkLabel(master=app, text=f"Checking SD-TAGs complete! \n"
                                                              f"Correct tags - {len(correct_tags)}\n"
                                                              f"Possible to fill SD - {len(possible_to_fill)}\n"
                                                              f"Waiting for load tags - {len(waiting_load_tag)}\n"
                                                              f"Waiting for sd tags - {len(waiting_sd_tag)}\n"
                                                              f"Wrong load tags - {len(wrong_load_tags)}\n"
                                                              f"Wrong sd tags - {len(wrong_sd_tags)}")
    result_checking.grid(row=5, column=1, pady=2, padx=0.5)

    print("correct tags - ", correct_tags)
    print("wrong load tags - ", wrong_load_tags)
    print("wrong sd tags - ", wrong_sd_tags)
    print("possible to fill tags - ", possible_to_fill)
    print("waiting load tags - ", waiting_load_tag)
    print("waiting sd tags - ", waiting_sd_tag)
    return
def start_checking():
    result_filling.destroy()
    progress_bar.destroy()

    btn_check_tags.configure(state=tkinter.DISABLED)
    thread = threading.Thread(target=check_load_sd_tags)
    print(threading.main_thread().name)
    print(thread.name)
    thread.start()
    check_thread(thread)
    pass

def fill_sd_tags():
    progress_bar = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    progress_bar.grid(row=5, column=1, pady=0.2, padx=0.5)
    result_filling = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    result_filling.grid(row=5, column=1, pady=2, padx=0.5)

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
                support_tag = "n/a"
                for attrib in entity.GetAttributes():
                    if "TAG" in attrib.TagString:
                        # print("!!!!!tag ", attrib.TextString)
                        support_tag = attrib.TextString.strip()
                    # update text
                    # if 'DRAWING' in attrib.TagString and "LATER" in attrib.TextString:
                    if 'DRAWING' in attrib.TagString:
                        # print(f"  --- {attrib.TextString}")
                        try:
                            if tag_list[support_tag] != "nan":
                                attrib.TextString = tag_list[support_tag]
                                attrib.Update()
                        except Exception as e:
                            print(e)
                            pass
    result_filling = customtkinter.CTkLabel(master=app, text=f"Filling SD-TAGs complete!")
    result_filling.grid(row=5, column=1, pady=2, padx=0.5)
    return
def start_fill():
    result_filling.destroy()
    progress_bar.destroy()

    btn_fill_tags.configure(state=tkinter.DISABLED)
    thread = threading.Thread(target=fill_sd_tags)
    print(threading.main_thread().name)
    print(thread.name)
    thread.start()
    check_thread(thread)
    return


def check_thread(thread):
    if thread.is_alive():
        app.after(100, lambda: check_thread(thread))
    else:
        btn_fill_tags.configure(state=tkinter.NORMAL)
        btn_check_tags.configure(state=tkinter.NORMAL)


btn_check_tags = customtkinter.CTkButton(master=app, command=start_checking, text="Check tags")
btn_check_tags.grid(row=4, column=1, pady=10, padx=0.1, sticky="w")

btn_fill_tags = customtkinter.CTkButton(master=app, command=start_fill, text="Fill tags")
btn_fill_tags.grid(row=4, column=1, pady=10, padx=210, sticky="w")






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