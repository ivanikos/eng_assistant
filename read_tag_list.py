import xlrd
import pandas as pd



# path_tag_list = (r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\cad_helper"
#                  r"\54880-13 Williams Pipe Support List.xls")
#
# path_tag_list_pd = (r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\cad_helper"
#                  r"\54880-13 Williams Pipe Support List.xls")


def read_tag_list(path_tag_list_pd):
    tag_list = {}
    try:
        wb = xlrd.open_workbook(path_tag_list_pd)
        ws = wb.sheet_by_name("Support Tags")
        for i in range(2000):
            project_number = ws.cell_value(i, 5)[2: 10]

            load_tag = ws.cell_value(i, 2)
            sd_tag = ws.cell_value(i, 11).replace(f"-{project_number}", "")

            tag_list[load_tag] = sd_tag
            print(project_number, load_tag, sd_tag)
    except Exception as e:
        print(f"ERROR - {e}")

    return tag_list

def read_tags_pd(path_tag_list_pd):
    tag_list = {}
    df = pd.read_excel(path_tag_list_pd, usecols=[2, 5, 11], sheet_name="Support Tags", header=7)

    for row in df.values:
        if row[0] != "nan" or "n":
            load_tag = row[0]
            project_number = str(row[1])[2: 10]
            sd_tag = str(row[2]).replace(f"-{project_number}", "")

            tag_list[load_tag] = sd_tag
            # print(project_number, load_tag, sd_tag)
        else:
            break
    return tag_list

# read_tags_pd()