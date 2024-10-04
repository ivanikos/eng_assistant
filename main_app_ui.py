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
from modules.logics import read_tags_pd

customtkinter.set_ctk_parent_class(tkinterdnd2.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("800x530")
app.title("Pipe Support Verifier v.0.02 (alpha testing)")


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

def checkbox_changed(*args):
    # This function will be called whenever the checkbox state changes
    if rev_circle_check_box.get():  # If the checkbox is checked
        print("Checkbox is checked")
        # Add logic for when the checkbox is checked
    else:
        print("Checkbox is unchecked")
        # Add logic for when the checkbox is unchecked

frame_1 = customtkinter.CTkFrame(master=app, corner_radius=10)
frame_1.grid(row=0, column=1, rowspan=4, sticky="nsew", pady=20, padx=20)

label_1 = customtkinter.CTkLabel(master=frame_1, text="Choose Pipe Support List for project:",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_1.grid(pady=10, padx=10, sticky="nsew")

browse_button = customtkinter.CTkButton(master=frame_1, fg_color="transparent", border_width=2,
                                                     text="Browse file", text_color=("gray10", "#DCE4EE"),
                                                     command=open_file)
browse_button.grid(row=2, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
file_name = customtkinter.CTkLabel(master=frame_1, text=f"{content}")
file_name.grid(row=2, column=0, pady=10, padx=1)

#
progress_bar = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
progress_bar.grid(row=5, column=1, pady=0.2, padx=0.5)
result_filling = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
result_filling.grid(row=6, column=1, pady=0.2, padx=0.5)


progress_bar_checking = customtkinter.CTkLabel(master=app, text=f"")
progress_bar_checking.grid(row=5, column=1, pady=0.2, padx=0.5)
result_checking = customtkinter.CTkLabel(master=app, text=f"")
result_checking.grid(row=6, column=1, pady=0.2, padx=0.5)

text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
text_1.grid(row=6, column=1, pady=2, padx=0.5)
text_1.insert("0.0", f"DETAILS:")

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

    text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    text_1.grid(row=6, column=1, pady=2, padx=0.5)
    text_1.insert("0.0", f"{fb.space_text}")

    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument  # Document object
    tag_list = read_tags_pd(content)

    print(acad.FullName) # full path to file

    print("Paperspace - ",len(acad.ActiveDocument.PaperSpace))

    progress = 0

    print("Blocks - ", len(doc.Blocks))

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
                            if load_tag == "LATER":
                                waiting_load_tag.append(f"{load_tag} - {sd_tag}")
                            elif load_tag in tag_list.keys() and tag_list[load_tag]:
                                if tag_list[load_tag] and sd_tag == "LATER":
                                    possible_to_fill.append(f"{load_tag} - {sd_tag} - in the support list - "
                                                                f"{tag_list[load_tag]}")
                                elif tag_list[load_tag] and sd_tag != "LATER":
                                    if tag_list[load_tag].strip() == sd_tag.strip():
                                        correct_tags.append(f"{load_tag} - {sd_tag} - correct")
                                    else:
                                        wrong_sd_tags.append(f"{load_tag} - {sd_tag} - INcorrect - in the support list "
                                                             f"- {tag_list[load_tag]}")
                                else:
                                    print(f"ELSE - {load_tag} - {sd_tag}")
                            elif load_tag in tag_list.keys() and not tag_list[load_tag]:
                                waiting_sd_tag.append(f"{load_tag} - {sd_tag} - wait sd-tag")
                            else:
                                wrong_load_tags.append(f"{load_tag} - {sd_tag} - need to double check LOAD TAG")
                        except Exception as e:
                            print(e)
                            # print(f"{attrib.TagString} -- {attrib.TextString} -- WRONG TAG")
        else:
            continue

    result_checking = customtkinter.CTkLabel(master=app, text=f"Checking SD-TAGs complete! \n"
                                                              f"Correct load + SD-tags - {len(correct_tags)}\n"
                                                              f"Possible to fill SD-tags - {len(possible_to_fill)}\n"
                                                              f"Waiting for load tags - {len(waiting_load_tag)}\n"
                                                              f"Waiting for SD-tags - {len(waiting_sd_tag)}\n"
                                                              f"Wrong load tags - {len(wrong_load_tags)}\n"
                                                              f"Wrong SD-tags - {len(wrong_sd_tags)}")
    result_checking.grid(row=5, column=1, pady=2, padx=0.5)


    # summarize detail result text
    detail_res = "DETAIL CHECKING RESULT:\n"

    if wrong_load_tags:
        detail_res += "\nWrong load tags: \n"
        for i in wrong_load_tags:
            detail_res = detail_res + f"{i}, \n"

    if wrong_sd_tags:
        detail_res += "\nWrong SD-tags: \n"
        for i in wrong_sd_tags:
            detail_res = detail_res + f"{i}, \n"

    # if correct_tags:
    #     detail_res += "\nCorrect tags: \n"
    #     for i in correct_tags:
    #         detail_res = detail_res + f"{i}, \n"

    if possible_to_fill:
        detail_res += "\nPossible to fill tags: \n"
        for i in possible_to_fill:
            detail_res = detail_res + f"{i}, \n"

    if waiting_load_tag:
        detail_res += "\nWaiting for load tags: \n"
        for i in waiting_load_tag:
            detail_res = detail_res + f"{i}, \n"

    if waiting_sd_tag:
        detail_res += "\nWaiting for sd tags: \n"
        for i in waiting_sd_tag:
            detail_res = detail_res + f"{i}, \n"




    # text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    # text_1.grid(row=6, column=1, pady=2, padx=0.5)
    text_1.insert("0.0", f"{detail_res}")

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
    text_1.destroy()

    btn_check_tags.configure(state=tkinter.DISABLED)
    thread = threading.Thread(target=check_load_sd_tags)
    print(threading.main_thread().name)
    print(thread.name)
    thread.start()
    check_thread(thread)
    return

def fill_sd_tags():
    filled_tags = []
    wrong_load_tags = []
    left_tags = []


    progress_bar = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    progress_bar.grid(row=5, column=1, pady=0.2, padx=0.5)
    result_filling = customtkinter.CTkLabel(master=app, text=f"{fb.space_text}")
    result_filling.grid(row=5, column=1, pady=2, padx=0.5)

    text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    text_1.grid(row=6, column=1, pady=2, padx=0.5)
    text_1.insert("0.0", f"{fb.space_text}")

    acad = win32com.client.Dispatch("AutoCAD.Application")

    doc = acad.ActiveDocument  # Document object
    tag_list = read_tags_pd(content)

    # text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    # text_1.grid(row=6, column=1, pady=2, padx=0.5)
    # text_1.insert("0.0", f"\n\n\n\n\n\n\n\n")

    progress = 0
    for entity in acad.ActiveDocument.PaperSpace:
        name = entity.EntityName

        progress_percentage = f"{round(((progress / len(acad.ActiveDocument.PaperSpace)) * 100), 0)} %"
        progress_bar = customtkinter.CTkLabel(master=app, text=f"{progress_percentage}")
        progress_bar.grid(row=5, column=1, pady=10, padx=0.5)

        progress += 1
        count_filled_tags = 0
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                load_tag = "n/a"
                for attrib in entity.GetAttributes():
                    if "TAG" in attrib.TagString:
                        # print("!!!!!tag ", attrib.TextString)
                        load_tag = attrib.TextString.strip()
                    # update text
                    # if 'DRAWING' in attrib.TagString and "LATER" in attrib.TextString:
                    if 'DRAWING' in attrib.TagString:
                        # print(f"  --- {attrib.TextString}")
                        if load_tag not in tag_list.keys():
                            wrong_load_tags.append(f"{load_tag} - wrong load_tag")
                        try:
                            if tag_list[load_tag] != "nan" and tag_list[load_tag]:
                                attrib.TextString = tag_list[load_tag]
                                attrib.Update()
                                filled_tags.append(f"{load_tag} - {tag_list[load_tag]} - OK")
                                count_filled_tags += 1
                            else:
                                left_tags.append(f"{load_tag} - waiting sd-tag")
                        except Exception as e:
                            print(e)
                            pass
    result_filling = customtkinter.CTkLabel(master=app, text=f"Filling SD-TAGs complete! \n"
                                                             f"filled - {len(filled_tags)} SD-tags.")
    result_filling.grid(row=5, column=1, pady=2, padx=0.5)

    # summarize detail result text
    detail_res = "DETAIL FILLING RESULT:\n"

    if left_tags:
        detail_res += "\nNOT filled tags: \n"
        for i in left_tags:
            detail_res = detail_res + f"{i}, \n"

    if wrong_load_tags:
        detail_res += "\nWrong load tags: \n"
        for i in wrong_load_tags:
            detail_res = detail_res + f"{i}, \n"

    # if filled_tags:
    #     detail_res += "\nFilled tags: \n"
    #     for i in filled_tags:
    #         detail_res = detail_res + f"{i}, \n"

    # text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    # text_1.grid(row=6, column=1, pady=2, padx=0.5)
    text_1.insert("0.0", f"{detail_res}")
    return
def start_fill():
    result_filling.destroy()
    progress_bar.destroy()
    text_1.destroy()

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
btn_check_tags.grid(row=4, column=1, pady=10, padx=50, sticky="w")

btn_fill_tags = customtkinter.CTkButton(master=app, command=start_fill, text="Fill tags")
btn_fill_tags.grid(row=4, column=1, pady=10, padx=260, sticky="w")

# Add a checkbox
rev_circle_check_box = customtkinter.BooleanVar(value=True)  # Variable to track checkbox state
checkbox = customtkinter.CTkCheckBox(app, text="Draw a marker on changed tags", variable=rev_circle_check_box)
checkbox.grid(row=4, column=1, pady=10, padx=450, sticky="w")
# Trace changes on the checkbox
rev_circle_check_box.trace("w", checkbox_changed)

print()




app.mainloop()