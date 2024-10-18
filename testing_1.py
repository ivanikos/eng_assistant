# import sys, io
#
# buffer = io.StringIO()
# sys.stdout = sys.stderr = buffer
import os
import tkinter
import tkinter.messagebox
import customtkinter
import tkinterdnd2
import threading
from tkinter.filedialog import askopenfile, asksaveasfilename
from CTkMessagebox import CTkMessagebox

import win32com.client

from modules import fill_block as fb
from modules.logics import read_tags_pd
from modules.logics import draw_marker
from modules.logics import export_report
from modules.import_data import help_message

customtkinter.set_ctk_parent_class(tkinterdnd2.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("600x560")
app.title("Pipe Support Verifier v.0.04 (alpha testing)")
app.grid_columnconfigure(1, weight=1)

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

def confirmation():
    # res = tkinter.messagebox.askquestion("Fill tags", "Do you really want to fill all SD-tags?")
    msg_box_confirmation = CTkMessagebox(title="Help", message="Do you really want to fill all tags?",
                                         option_1="Cancel", option_2="Yes")
    user_response = msg_box_confirmation.get()
    if user_response == "Yes":
        return 1
    else:
        return 0

def help_menu_action():
    # Create a new top-level window for help
    help_window = customtkinter.CTkToplevel(app)
    help_window.title("Help")
    help_window.geometry("530x240")  # Set the desired window size
    help_window.grid_columnconfigure(1, weight=1)
    help_window.resizable(False, False)
    # help_window.focus_force()
    # Help message content
    help_window.attributes('-topmost', True)

    help_message = (
        "For detailed instructions on using the application, please refer to the User Guide.\n\n\n"
        "If you encounter any issues or have feedback, please contact us at:\n\n"
        "ivanignatenko@uccenvironmental.com"
    )

    frame_1 = customtkinter.CTkFrame(master=help_window, corner_radius=10)
    frame_1.grid(row=0, column=1, rowspan=4, sticky="nsew", pady=10, padx=10)

    label_1 = customtkinter.CTkLabel(master=frame_1, text="Pipe Support Verifier v0.04 (alpha test)",
                                     font=customtkinter.CTkFont(size=20, weight="bold"))
    label_1.grid(pady=10, padx=10, sticky="nsew")

    label_2 = customtkinter.CTkLabel(master=frame_1, text=f"{help_message}",
                                     font=customtkinter.CTkFont(size=14))
    label_2.grid(pady=10, padx=10, sticky="nsew")

    guide_path = r"N:\Piping\_H_PPSE Users\IvaIgn\PSV\Pipe Support Verifier User Guide.pdf"

    def open_user_guide():
        # Open the user guide PDF
        try:
            if os.path.exists(guide_path):
                os.startfile(guide_path)
            else:
                print("User Guide not found!")
        except Exception as e:
            print(f"Error opening user guide: {e}")

    def open_email():
        # Open the default email client with a new message
        email_address = "ivanignatenko@uccenvironmental.com"
        try:
            os.startfile(f"mailto:{email_address}")
        except Exception as e:
            print(f"Error opening email client: {e}")

    # Create action buttons
    contact_button = customtkinter.CTkButton(frame_1, text="Contact Us", command=open_email)
    contact_button.grid(row=4, column=0, padx=10, pady=5, sticky="s")

    guide_button = customtkinter.CTkButton(frame_1, text="Open User Guide", command=open_user_guide)
    guide_button.grid(row=4, column=0, padx=10, pady=5, sticky="w")

    close_button = customtkinter.CTkButton(frame_1, text="Close", command=help_window.destroy)
    close_button.grid(row=4, column=0, padx=10, pady=5, sticky="e")

    return


report_to_export = []




user_name = os.getlogin()
frame_1 = customtkinter.CTkFrame(master=app, corner_radius=10)
frame_1.grid(row=0, column=1, rowspan=4, sticky="nsew", pady=10, padx=10)

label_1 = customtkinter.CTkLabel(master=frame_1, text="Choose Pipe Support List for project:",
                                    font=customtkinter.CTkFont(size=20, weight="bold"))
label_1.grid(pady=10, padx=10, sticky="nsew")

browse_button = customtkinter.CTkButton(master=frame_1, fg_color="transparent", border_width=2,
                                                     text="Browse file", text_color=("gray10", "#DCE4EE"),
                                                     command=open_file)
browse_button.grid(row=2, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
file_name = customtkinter.CTkLabel(master=frame_1, text=f"{content}")
file_name.grid(row=2, column=0, pady=10, padx=1)


frame_2 = customtkinter.CTkFrame(master=app, corner_radius=10)
frame_2.grid(row=4, column=1, rowspan=3, sticky="nsew", pady=10, padx=10)
button_definition_label = customtkinter.CTkLabel(master=frame_2, justify='left', text='"Check tags" button - check if                     "Fill tags" button is available\n'
                                                                  'load tags match SD-tags in the                    only after using the "Checking" feature\n'
                                                                  'pipe support list')
button_definition_label.grid(row=4, column=0, pady=10, padx=30, sticky="w")

progress_bar = customtkinter.CTkProgressBar(master=app, width=200)
progress_bar.set(0)
progress_bar.grid_forget()

# Create a Menu Bar (from tkinter)
menu_bar = tkinter.Menu(app)
file_menu = tkinter.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Placeholder", command=lambda: print("Placeholder"))
file_menu.add_separator()
file_menu.add_command(label="Exit", command=app.quit)

# Add the file menu to the menu bar
menu_bar.add_cascade(label="File", menu=file_menu)

# Create another menu for Edit options
help_menu = tkinter.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Help", command=help_menu_action)
help_menu.add_separator()
menu_bar.add_cascade(label="Help", menu=help_menu)

# Add the menu bar to the app window
app.config(menu=menu_bar)


def export():
    text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    text_1.grid(row=7, column=1, pady=10, padx=10, sticky="n")
    text_1.insert("0.0", f"{fb.space_text}")
    try:
        # Open the save file dialog
        file_path = asksaveasfilename(
            defaultextension=".txt",  # Default file extension
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            export_report(report_to_export, file_path)
            print(f"Path to report - {file_path}")
        else:
            print("NO filename")
    except Exception as e:
        text_1.insert("0.0", f"ERROR: \n {e}")

def check_load_sd_tags():
    global report_to_export
    report_to_export = []

    all_load_tags = []
    correct_tags = []
    wrong_load_tags = []
    wrong_sd_tags = []
    possible_to_fill = []
    waiting_load_tag = []
    waiting_sd_tag = []
    duplicated_load_tags = {}

    text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    text_1.grid(row=7, column=1, pady=10, padx=10, sticky="nsew")
    text_1.insert("0.0", f"{fb.space_text}")

    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument  # Document object
        doc_filename = doc.FullName
        tag_list = read_tags_pd(content)

        progress_bar.grid(row=8, column=1, pady=10, padx=10)
        progress_bar.set(0)
        progress = 0
        full_progress = len(acad.ActiveDocument.PaperSpace) - 1
        all_checked_tags = 0

        print(doc_filename)
        print("Blocks - ", len(doc.Blocks))
        print("Paperspace - ",len(acad.ActiveDocument.PaperSpace))

        for entity in acad.ActiveDocument.PaperSpace:
            name = entity.EntityName

            progress_percentage = progress / full_progress
            progress_bar.set(progress_percentage)
            app.update_idletasks()

            progress += 1
            if name == 'AcDbBlockReference':
                HasAttributes = entity.HasAttributes
                if HasAttributes:
                    load_tag = "n/a"
                    for attrib in entity.GetAttributes():
                        if "DWGNO" in attrib.TagString:
                            print(attrib.TextString.strip())

                        if "TAG" in attrib.TagString:
                            load_tag = attrib.TextString.strip()
                            if load_tag.upper() != "LATER":
                                all_load_tags.append(load_tag.strip())

                        if "DRAWING" in attrib.TagString and attrib.TextString:
                            # print(load_tag, attrib.TextString)
                            all_checked_tags += 1
                            sd_tag = str(attrib.TextString).strip()
                            try:
                                if load_tag.upper() == "LATER":
                                    waiting_load_tag.append(f"{load_tag} - {sd_tag}")
                                elif load_tag in tag_list.keys() and tag_list[load_tag]:
                                    if tag_list[load_tag] and sd_tag.upper() == "LATER":
                                        possible_to_fill.append(f"{load_tag} - {sd_tag} - in the support list - "
                                                                    f"{tag_list[load_tag]}")
                                    elif tag_list[load_tag] and sd_tag.upper() != "LATER":
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
                                    wrong_load_tags.append(f"{load_tag.strip().replace("-", "n/a")} - {sd_tag} - need to double check LOAD TAG")
                            except Exception as e:
                                print(e)
            else:
                continue

        # checking for duplicates
        for tag in all_load_tags:
            if all_load_tags.count(tag) > 1:
                duplicated_load_tags[tag] = all_load_tags.count(tag)

        # preparing to export report
        report_to_export.append([f"Working file name: \n {doc_filename}", "", ""])
        report_to_export.append(["***", "***", "***"])
        report_to_export.append([f"Checking SD-TAGs complete!", "", ""])
        report_to_export.append([f"Checked tags - {all_checked_tags}", "", ""])
        report_to_export.append([f"Correct load + SD-tags - {len(correct_tags)}", "", ""])
        report_to_export.append([f"Possible to fill SD-tags - {len(possible_to_fill)}", "", ""])
        report_to_export.append([f"Waiting for load tags - {len(waiting_load_tag)}", "", ""])
        report_to_export.append([f"Waiting for SD-tags - {len(waiting_sd_tag)}", "", ""])
        report_to_export.append([f"Wrong load tags - {len(wrong_load_tags)}", "", ""])
        report_to_export.append([f"Wrong SD-tags - {len(wrong_sd_tags)} ", "", ""])
        report_to_export.append(["***", "***", "***"])
        report_to_export.append(["DETAIL CHECKING RESULT:", "", ""])

        # summarize detail result text
        result_overview = f"Working file name: \n {doc_filename}\n\n" \
                f"Checking SD-TAGs complete! \n" \
                f"Checked tags - {all_checked_tags}\n" \
                f"Correct load + SD-tags - {len(correct_tags)}\n" \
                f"Possible to fill SD-tags - {len(possible_to_fill)}\n" \
                f"Waiting for load tags - {len(waiting_load_tag)}\n" \
                f"Waiting for SD-tags - {len(waiting_sd_tag)}\n" \
                f"Wrong load tags - {len(wrong_load_tags)}\n" \
                f"Duplicated load-tags - {len(duplicated_load_tags.keys())} \n" \
                f"Wrong SD-tags - {len(wrong_sd_tags)} \n\n" \


        detail_res = result_overview + "DETAIL CHECKING RESULT:\n"

        if wrong_load_tags:
            detail_res += "\nLoad tags does not exist: \n"
            report_to_export.append(["Load tags does not exist:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", ""])
            for i in wrong_load_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.replace(" - need to double check LOAD TAG", "").split("-")[0]
                sd_t = i.replace(" - need to double check LOAD TAG", "").split("-")[1]
                report_to_export.append([l_t, sd_t, ""])
            report_to_export.append(["***", "***", "***"])

        if wrong_sd_tags:
            detail_res += "\nWrong SD-tags: \n"
            report_to_export.append(["SD tags does not match:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", "SD tag in pipe support list"])
            for i in wrong_sd_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.replace("- INcorrect - in the support list", "") \
                    .replace(" - ", "&").split("&")[0]
                sd_t = i.replace("- INcorrect - in the support list", "") \
                    .replace(" - ", "&").split("&")[1]
                csd_t = i.replace("- INcorrect - in the support list", "") \
                    .replace(" - ", "&").split("&")[2]
                report_to_export.append([l_t, sd_t, csd_t])
            report_to_export.append(["***", "***", "***"])

        if duplicated_load_tags.keys():
            detail_res += "\nDuplicated load-tags: \n"
            report_to_export.append(["Load tags exist more than one time:", "", ""])
            report_to_export.append(["Actual load tag", "Number of duplicates", ""])
            for i in duplicated_load_tags.keys():
                detail_res = detail_res + f"Load tag - {i} - number of duplicates - {duplicated_load_tags[i]}, \n"
                l_t = i
                sd_t = duplicated_load_tags[i]
                csd_t = ""
                report_to_export.append([l_t, sd_t, csd_t])
            report_to_export.append(["***", "***", "***"])

        if possible_to_fill:
            detail_res += "\nPossible to fill tags: \n"
            report_to_export.append(["Tags ready to fill:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", "SD tag in pipe support list"])
            for i in possible_to_fill:
                detail_res = detail_res + f"{i}, \n"
                # .replace("- in the support list", "").replace(" - ", "&").split("&"))
                l_t = i.replace("- in the support list", "").replace(" - ", "&").split("&")[0]
                sd_t = i.replace("- in the support list", "").replace(" - ", "&").split("&")[1]
                csd_t = i.replace("- in the support list", "").replace(" - ", "&").split("&")[2]
                report_to_export.append([l_t, sd_t, csd_t])
            report_to_export.append(["***", "***", "***"])

        if waiting_load_tag:
            detail_res += "\nWaiting for load tags: \n"
            report_to_export.append(["Load tag - LATER:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", ""])
            for i in waiting_load_tag:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.split("-")[0]
                sd_t = i.split("-")[1]
                report_to_export.append([l_t, sd_t, ""])
            report_to_export.append(["***", "***", "***"])

        if waiting_sd_tag:
            detail_res += "\nWaiting for sd tags: \n"
            report_to_export.append(["SD tag does not exist yet:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", ""])
            for i in waiting_sd_tag:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.replace(" - ", "&").split("&")[0]
                sd_t = i.replace(" - ", "&").split("&")[1]
                report_to_export.append([l_t, sd_t, ""])
            report_to_export.append(["***", "***", "***"])

        if correct_tags:
            detail_res += "\nCorrect tags: \n"
            report_to_export.append(["Correct load + SD tags:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", ""])
            for i in correct_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.split(" - ")[0]
                sd_t = i.split(" - ")[1]
                report_to_export.append([l_t, sd_t, ""])
            report_to_export.append(["***", "***", "***"])



        text_1.insert("0.0", f"{detail_res}")
        progress_bar.grid_forget()
        btn_export_report.configure(state=tkinter.NORMAL)
        btn_fill_tags.configure(state=tkinter.NORMAL)

    except Exception as e:
        text_1.insert("0.0", f"ERROR: \n {e}")
        progress_bar.grid_forget()



    # print("correct tags - ", correct_tags)
    # print("wrong load tags - ", wrong_load_tags)
    # print("wrong sd tags - ", wrong_sd_tags)
    # print("possible to fill tags - ", possible_to_fill)
    # print("waiting load tags - ", waiting_load_tag)
    # print("waiting sd tags - ", waiting_sd_tag)

    return
def start_checking():
    text_1.destroy()
    btn_check_tags.configure(state=tkinter.DISABLED)
    btn_fill_tags.configure(state=tkinter.DISABLED)
    btn_export_report.configure(state=tkinter.DISABLED)
    thread = threading.Thread(target=check_load_sd_tags)
    print(threading.main_thread().name)
    print(thread.name)
    thread.start()
    check_thread(thread)
    return

def fill_sd_tags():
    global report_to_export
    report_to_export = []

    filled_tags = []
    wrong_load_tags = []
    left_tags = []

    draw_check_box = rev_circle_check_box.get()

    text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
    text_1.grid(row=7, column=1, pady=10, padx=10, sticky="n")
    text_1.insert("0.0", f"{fb.space_text}")

    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")

        doc = acad.ActiveDocument  # Document object
        doc_filename = doc.FullName
        tag_list = read_tags_pd(content)

        progress_bar.grid(row=8, column=1, pady=10, padx=10)
        progress_bar.set(0)
        progress = 0
        count_filled_tags = 0
        count_not_filled_tags = 0

        full_progress = len(acad.ActiveDocument.PaperSpace) - 1

        for entity in acad.ActiveDocument.PaperSpace:
            name = entity.EntityName

            progress_percentage = progress / full_progress
            progress_bar.set(progress_percentage)
            app.update_idletasks()

            progress += 1
            if name == 'AcDbBlockReference':
                HasAttributes = entity.HasAttributes
                insertion_point = entity.InsertionPoint
                if HasAttributes:

                    # Getting scale of block
                    scale_x = entity.XScaleFactor
                    scale_y = entity.YScaleFactor
                    scale_z = entity.ZScaleFactor
                    coord_list = [float(insertion_point[0]), float(insertion_point[1]), float(insertion_point[2])]

                    load_tag = "n/a"
                    for attrib in entity.GetAttributes():
                        if "TAG" in attrib.TagString:
                            load_tag = attrib.TextString.strip()
                        if 'DRAWING' in attrib.TagString:
                            if load_tag not in tag_list.keys():
                                wrong_load_tags.append(f"{load_tag} - load_tag does not exist")
                                count_not_filled_tags += 1
                            try:
                                if tag_list[load_tag] != "nan" and tag_list[load_tag]:
                                    attrib.TextString = tag_list[load_tag]
                                    attrib.Update()

                                    if draw_check_box:
                                        draw_marker(coord_list, 0.2 * float(scale_x), color=4)

                                    filled_tags.append(f"{load_tag} - {tag_list[load_tag]} - OK")
                                    count_filled_tags += 1
                                else:
                                    left_tags.append(f"{load_tag} - waiting sd-tag")
                                    count_not_filled_tags += 1
                            except Exception as e:
                                print(e)
                                pass

        # preparing to export report
        report_to_export.append([f"Working file name: \n {doc_filename}", "", ""])
        report_to_export.append(["***", "***", "***"])
        report_to_export.append([f"Filling SD-TAGs complete!", "", ""])
        report_to_export.append([f"Filled - {count_filled_tags} SD-tags.", "", ""])
        report_to_export.append([f"NOT Filled - {count_not_filled_tags} SD-tags.", "", ""])
        report_to_export.append(["***", "***", "***"])
        report_to_export.append(["DETAIL CHECKING RESULT:", "", ""])


        # summarize detail result text
        result_text = f"Working file name: \n {doc_filename}\n\n" \
                       f"Filling SD-TAGs complete! \n" \
                        f"Filled - {count_filled_tags} SD-tags.\n"
        detail_res = result_text + "\nDETAIL FILLING RESULT:\n"

        if left_tags:
            detail_res += "\nNOT filled tags: \n"
            report_to_export.append(["NOT filled tags - waiting SD-tag:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", ""])
            for i in left_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.split(" - ")[0]
                report_to_export.append([l_t, "Waiting SD-tag", ""])
            report_to_export.append(["***", "***", "***"])

        if wrong_load_tags:
            detail_res += "\nLoad tags doesn't match support list: \n"
            report_to_export.append(["Load tags doesn't exist in support list:", "", ""])
            report_to_export.append(["Actual load tag", "Existing status", ""])
            for i in wrong_load_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.split(" - ")[0]
                report_to_export.append([l_t, "Check load-tag", ""])
            report_to_export.append(["***", "***", "***"])

        if filled_tags:
            detail_res += "\nFilled tags: \n"
            report_to_export.append(["Filled tags:", "", ""])
            report_to_export.append(["Actual load tag", "Actual SD tag", "Status"])
            for i in filled_tags:
                detail_res = detail_res + f"{i}, \n"
                l_t = i.split(" - ")[0]
                sd_t = i.split(" - ")[1]
                status = i.split(" - ")[2]
                report_to_export.append([l_t, sd_t, status])
            report_to_export.append(["***", "***", "***"])

        text_1.insert("0.0", f"{detail_res}")
        progress_bar.grid_forget()
        btn_export_report.configure(state=tkinter.NORMAL)
        btn_fill_tags.configure(state=tkinter.NORMAL)

    except Exception as e:
        text_1.insert("0.0", f"{e}")
        progress_bar.grid_forget()


    return
def start_fill():
    text_1.destroy()
    confirmation_res = confirmation()
    if confirmation_res == 1:
        btn_fill_tags.configure(state=tkinter.DISABLED)
        btn_check_tags.configure(state=tkinter.DISABLED)
        btn_export_report.configure(state=tkinter.DISABLED)
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
        btn_check_tags.configure(state=tkinter.NORMAL)



btn_export_report = customtkinter.CTkButton(master=app, command=export, text="Export to .xlsx")
# btn_export_report = customtkinter.CTkButton(master=app, text="Export to .xlsx")
btn_export_report.grid(row=8, column=1, padx=30, pady=10, sticky="nw")
btn_export_report.configure(state=tkinter.DISABLED)

user_name_ui = customtkinter.CTkLabel(master=app, text=f"User: {user_name}",
                                      font=customtkinter.CTkFont(size=10, weight="bold"))
user_name_ui.grid(row=8, column=1, padx=30, pady=10, sticky="se")

btn_check_tags = customtkinter.CTkButton(master=frame_2, command=start_checking, text="Check tags")
btn_check_tags.grid(row=5, column=0, padx=30, pady=10, sticky="w")

btn_fill_tags = customtkinter.CTkButton(master=frame_2, command=start_fill, text="Fill tags")
btn_fill_tags.grid(row=5, column=0, padx=250, pady=10, sticky="w")
btn_fill_tags.configure(state=tkinter.DISABLED)

# Add a checkbox
rev_circle_check_box = customtkinter.BooleanVar(value=True)  # Variable to track checkbox state
checkbox = customtkinter.CTkCheckBox(master=frame_2, text="Draw a marker on changed tags", variable=rev_circle_check_box)
checkbox.grid(row=6, column=0, pady=10, padx=250, sticky="w")

text_1 = customtkinter.CTkTextbox(master=app, width=600, height=170)
text_1.grid(row=7, column=1, pady=10, padx=10, sticky="nsew")
text_1.insert("0.0", f"DETAILS:")




app.mainloop()

