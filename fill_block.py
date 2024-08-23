import win32com.client
from read_tag_list import read_tag_list, read_tags_pd


# tag_list = {"C100310": "SD-111", "C100300": "SD-321", "C100290": "SD-222", "n/a": "LATER1", "P800100": "success"}
# tag_list = read_tag_list()
# tag_list = read_tags_pd()


# iterate through all objects (entities) in the currently opened drawing
# and if it's a BlockReference, display its attributes and some other things.

def fill_support_tags(path):
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
                    print("******")
                    print("  {}: {}".format(attrib.TagString, attrib.TextString))
                    print("******")

                    if "TAG" in attrib.TagString:
                        print("!!!!!tag ", attrib.TextString)
                        support_tag = attrib.TextString.strip()
                        # print("sd ", support_tag)
                    # update text
                    # if 'DRAWING' in attrib.TagString and "LATER" in attrib.TextString:
                    if 'DRAWING' in attrib.TagString:
                        print(f"  --- {attrib.TextString}")
                        try:
                            attrib.TextString = tag_list[support_tag]
                            attrib.Update()
                        except Exception as e:
                            print(e)
                        continue

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
                                if str(attrib.TextString).strip() ==  tag_list[support_tag].strip():
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


                        continue










