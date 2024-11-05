import os
from datetime import datetime

import pandas as pd
from pyautocad import Autocad, APoint
import xlsxwriter


"""Read the tag-list .xlsm/xlsx etc."""
def read_tags_pd(path_tag_list_pd):
    tag_list = {}

    df_cn = pd.read_excel(path_tag_list_pd, usecols=[2, 4], sheet_name="Support Tags", header=0)
    contract_number = "N/A"
    for row in df_cn.values:
        if (row[0] != "nan" or "n") and row[0] and r"N/A" in contract_number:
            if "Contract Number:" in row[0]:
                contract_number = row[1]
                print(contract_number)

    df = pd.read_excel(path_tag_list_pd, usecols=[2, 5, 11], sheet_name="Support Tags", header=7)
    for row in df.values:
        if row[0] != "nan" or "n":
            load_tag = row[0]
            # contract_number = str(row[1])[2: 10]
            sd_tag = str(row[2]).replace(f"-{contract_number}", "").replace("nan", "")

            tag_list[load_tag] = sd_tag

        else:
            break
    return tag_list


def draw_marker(center, radius, color=1):
    acad_1 = Autocad(create_if_not_exists=True)
    # Access the active document
    doc_1 = acad_1.ActiveDocument

    center_point = APoint(center[0], center[1], center[2])
    circle = doc_1.PaperSpace.AddCircle(center_point, radius)
    circle.Color = color


def export_report(report, filename):
    workbook_summary = xlsxwriter.Workbook(f'{filename}')
    # -------------------------------------Краткая сводка
    ws0 = workbook_summary.add_worksheet('Report')
    ws0.set_column(0, 2, 25)
    ws0.set_row(0, 45)
    # ws0.set_column(1, 1, 40)
    # ws0.set_column(4, 15, 12)
    # ws0.set_column(2, 2, 12)
    # ws0.set_column(3, 3, 12)

    cell_format_green = workbook_summary.add_format()
    cell_format_green.set_bg_color('#6c8784')
    cell_format_blue = workbook_summary.add_format()
    cell_format_blue.set_bg_color('#b0dacc')
    cell_format_hat = workbook_summary.add_format()
    cell_format_hat.set_bg_color('#65878f')
    cell_format_filename = workbook_summary.add_format()
    cell_format_filename.set_bg_color('#f8e8b5')


    for i, (one, two, three) in enumerate(report, start=1):
        color = cell_format_blue
        if "Working" in one:
            color = cell_format_filename
            color.set_text_wrap(text_wrap=1)
            ws0.merge_range('A1:C1', one, color)
            continue
        if "***" in one:
            color = cell_format_green

        if "Checking SD-TAGs complete" in one or \
            "DETAIL CHECKING" in one or "Load tags does" in one or "SD tag" in one:
            color.set_text_wrap(text_wrap=1)
            color = cell_format_hat
        try:
            color.set_border(style=1)
            color.set_text_wrap(text_wrap=1)
        except:
            pass

        ws0.write(f'A{i}', one, color)
        ws0.write(f'B{i}', two, color)
        ws0.write(f'C{i}', three, color)

    workbook_summary.close()
    print("Success")
    return

def write_log(user_name, message):
    log_path = r"N:\Piping\_H_PPSE Users\IvaIgn\PSV\logs\\"

    log_file_path = f"{log_path}{user_name}_log.txt"
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log_entry = f"{current_time} - {message}\n"

    with open(log_file_path, "a") as log_file:
        log_file.write(log_entry)