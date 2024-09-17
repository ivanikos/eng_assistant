from itertools import count

import win32com.client
import pyautocad
from comtypes.client import *
from comtypes.automation import *


# comtypes.automation._vartype_to_ctype[9] = comtypes.automation.VARIANT
from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)
doc = acad.ActiveDocument

# acad = win32com.client.Dispatch("AutoCAD.Application")
# doc = acad.ActiveDocument  # Document object

print(doc.name)
needed_blocks = []


for block in acad.iter_objects('AcDbBlockReference', dont_cast=True):
    print(block.Name)
    print(block.InsertionPoint)
    print(block.Layer)

    if block.Layer == "SUPPORT_TAGS":
        print(block.Name, "!!!")
        # print(block.InsertionPoint)
        # print(block.ObjectName)
        # print(block.ObjectID)
        # print(block.Handle)
        needed_blocks.append(block.Handle)
    else:
        continue

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument  # Document object

print(len(needed_blocks))
for obj in needed_blocks:
    object = acad.ActiveDocument.HandleToObject(obj)
    HasAttributes = object.HasAttributes
    if HasAttributes:
        for attrib in object.GetAttributes():
            print(attrib.TagString)
            print(attrib.TextString)



# for attr in obj.GetAttributes():
#     print(attr.TagString)
#     print(attr.TextString)