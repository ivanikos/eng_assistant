import os
import tkinter
import tkinter.messagebox
import customtkinter
import tkinterdnd2
import threading
from tkinter.filedialog import askopenfile, asksaveasfilename
from CTkMessagebox import CTkMessagebox
from CTkMenuBar import *

import win32com.client
from pyautocad import Autocad, APoint



# acad = win32com.client.Dispatch("AutoCAD.Application")

# acad = win32com.client.GetObject(None, "AutoCAD.Application")

try:
    # Try to get an existing AutoCAD instance
    acad = win32com.client.GetObject(None, "AutoCAD.Application")

except Exception:
    # AutoCAD isn't running
    print("AutoCAD is not open.")


print(acad.visible)
print(acad.Documents.Count)
if acad.Documents.Count == 0:
    print(" No active drawings are open")

doc = acad.ActiveDocument  # Document object
doc_filename = doc.FullName

print(doc_filename)
print("Blocks - ", len(doc.Blocks))
print("Paperspace - ",len(acad.ActiveDocument.PaperSpace))

dwg_sheets = {}

for entity in acad.ActiveDocument.PaperSpace:
    name = entity.EntityName

    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        if HasAttributes:
            load_tag = "n/a"

            insertion_point = entity.InsertionPoint
            scale_x = entity.XScaleFactor
            scale_y = entity.YScaleFactor
            scale_z = entity.ZScaleFactor
            coord_list = [float(insertion_point[0]), float(insertion_point[1]), float(insertion_point[2])]

            center_point = APoint(coord_list[0], coord_list[1], coord_list[2])
            circle = doc.PaperSpace.AddCircle(center_point, 5)

            for attrib in entity.GetAttributes():
                if attrib.TagString.strip() == "DWGNO":
                    dwg_sheets[attrib.TextString.strip()] = [entity.Handle, scale_x, coord_list]
                    # print(f"HANDLE - {entity.Handle}, TEXT STRING - {attrib.TextString.strip()}, "
                    #       f"TAG STRING - {attrib.TagString.strip()}, COORD LIST - {coord_list}, SCALE - {scale_x}")

    else:
        continue

for i in dwg_sheets:
    print(f"DWG - {i}, HANDLE - {dwg_sheets[i][0]}, SCALE - {dwg_sheets[i][1]}, COORDS - {dwg_sheets[i][2]}")


del acad