import win32com.client
import pandas as pd
# from modules import read_tag_list as rtl


# tag_list = {"C100310": "SD-111", "C100300": "SD-321", "C100290": "SD-222", "n/a": "LATER1", "P800100": "success"}
# tag_list = read_tag_list()
# tag_list = read_tags_pd()


# iterate through all objects (entities) in the currently opened drawing
# and if it's a BlockReference, display its attributes and some other things.

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

def fill_support_tags(path):
    acad = win32com.client.Dispatch("AutoCAD.Application")

    doc = acad.ActiveDocument  # Document object

    tag_list = read_tags_pd(path)

    progress = 0
    for entity in acad.ActiveDocument.PaperSpace:
        name = entity.EntityName
        # print(f"{round(((progress / len(acad.ActiveDocument.PaperSpace)) * 100), 2)} %")
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
    return


def check_sd_tags(path):
    acad = win32com.client.Dispatch("AutoCAD.Application")

    doc = acad.ActiveDocument  # Document object

    tag_list = read_tags_pd(path)
    for entity in acad.ActiveDocument.PaperSpace:
        name = entity.EntityName
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
                        # print("TAG ", attrib.TextString)
                        support_tag = attrib.TextString.strip()

                    if "DRAWING" in attrib.TagString and attrib.TextString:
                        try:
                            if tag_list[support_tag] and attrib.TextString == "LATER":
                                print(f"{attrib.TextString, tag_list[support_tag]}  - possible to fill")
                            else:
                                if str(attrib.TextString).strip() == tag_list[support_tag].strip():
                                    # print(f"Load tag - {attrib.TagString} \n "
                                    #       f"SD-tag - {tag_list[support_tag]} \n"
                                    #       f"OK \n")
                                    pass
                                else:
                                    print(f"Load tag - {support_tag} \n "
                                          f"SD-tag - {str(attrib.TextString).strip()} \n"
                                          f"INCORRECT. Should be - {tag_list[support_tag]}\n")

                        except:
                            print(f"{attrib.TagString} -- {attrib.TextString} -- WRONG TAG")
                            pass

    return

space_text = ("                                                                       \n"
              "                                                                       \n"
              "                                                                       \n"
              "                                                                       \n"
              "                                                                       \n"
              "                                                                       \n"
              "                                                                       \n")

# fill_support_tags(r"C:/Users/vanik/PycharmProjects/cad_helper/54880-13 Williams Pipe Support List.xlsm")
# fill_support_tags(r"C:/Users/ivaign/OneDrive - United Conveyor Corp/Documents/Python_Projects/cad_helper/54880-13 Williams Pipe Support List.xlsm")
