import win32com.client
from comtypes.client import *
from comtypes.automation import *
import win32com.client
import numpy



# acad = win32com.client.DispatchEx("AutoCAD.Application")
acad = comtypes.client.GetActiveObject('AutoCAD.Application', dynamic=True)

doc = acad.ActiveDocument  # Document object
acad.Visible = True

doc_filename = doc.FullName

print(doc.Name)
print(acad.Documents.Count)

del acad