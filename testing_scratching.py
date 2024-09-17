from itertools import count

import win32com.client
import pyautocad
from comtypes.client import *
from comtypes.automation import *
from fontTools.unicodedata import block

# comtypes.automation._vartype_to_ctype[9] = comtypes.automation.VARIANT
from pyautocad import Autocad

# acad = Autocad(create_if_not_exists=True)
# doc = acad.ActiveDocument
# print(doc.name)
#
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument  # Document object
print(doc.name)

needed_blocks = []

# print("Paperspace - ",len(doc.PaperSpace))
# print("blocks - ", len(doc.Blocks))
process = 0

def get_id():
    acad = Autocad(create_if_not_exists=True)
    doc = acad.ActiveDocument

    for block in acad.iter_objects('AcDbBlockReference', dont_cast=True):
    # process = len(doc.Blocks)
    # for block in doc.Blocks:
        # print(block.InsertionPoint)
        # print(block.Layer)
        print("NAME - ", block.Name)
        print("Handle - ", block.Handle)
        print("ObjectID - ", block.ObjectID)
        # print(dir(block))

        needed_blocks.append(block.Handle)

        # process += 1
        # print(process)


        # try:
        #     print(block.Layer)
        #     print(block.Name)
        # except:
        #     pass

        #
        # if block.Layer == "SUPPORT_TAGS":
        #     # print(block.Name, "!!!")
        #     # print(block.InsertionPoint)
        #     # print(block.ObjectName)
        #     # print(block.ObjectID)
        #     print(block.Handle)
        #     needed_blocks.append(block.Handle)
        # else:
        #     continue
    return needed_blocks
get_id()




# acad = win32com.client.Dispatch("AutoCAD.Application")
# doc = acad.ActiveDocument  # Document object
# print(doc.name)
#
# needed_blocks = get_id()
# print("Needed block list - ", len(needed_blocks))
#
# for obj in needed_blocks:
#     object = acad.ActiveDocument.HandleToObject(obj)
#
#     HasAttributes = object.HasAttributes
#     if HasAttributes:
#         for attrib in object.GetAttributes():
#             print(attrib.TagString)
#             print(attrib.TextString)





