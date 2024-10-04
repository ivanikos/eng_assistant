from pyautocad import Autocad, APoint
import win32com.client
import math
from array import array



acad = win32com.client.Dispatch("AutoCAD.Application")
print(acad.FullName)



doc = acad.ActiveDocument  # Document object
space = doc.PaperSpace
print("Paperspace - ",len(acad.ActiveDocument.PaperSpace))
print("Blocks - ", len(doc.Blocks))
print(doc.FullName)
print(doc.Name)


coord_scale_1 = [980.4079781921173, -2171.1298871479075, 0.0]



# Function to create a circular cloud shape around a given point
def draw_cloud(center, radius, num_bumps):
    angle_step = 2 * math.pi / num_bumps
    points = []

    acad_1 = Autocad(create_if_not_exists=True)

    # Access the active document
    doc_1 = acad_1.ActiveDocument


    # Calculate the points for each bump on the cloud
    for i in range(num_bumps + 1):  # +1 to close the loop
        angle = i * angle_step
        x = center[0] + radius * math.cos(angle)
        y = center[1] + radius * math.sin(angle)
        z = center[2]  # Keep the Z coordinate constant
        points.extend([x, y, z])  # Flatten the point coordinates

    # Create a variant array for the polyline points
    polyline_points = array('d', points)  # 'd' stands for double (floating point)

    # Create a polyline in AutoCAD to form the cloud
    polyline = doc_1.PaperSpace.AddPolyline(polyline_points)
    polyline.Closed = True  # Close the polyline to make it a loop

def draw_circle(center, radius, color=1):
    acad_1 = Autocad(create_if_not_exists=True)
    # Access the active document
    doc_1 = acad_1.ActiveDocument

    center_point = APoint(center[0], center[1], center[2])
    circle = doc_1.PaperSpace.AddCircle(center_point, radius)
    circle.Color = color


for entity in acad.ActiveDocument.PaperSpace:
    name = entity.EntityName

    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        insertion_point = entity.InsertionPoint

        coord_list = [float(insertion_point[0]), float(insertion_point[1]), float(insertion_point[2])]

        # center = (abs(float(insertion_point[0])), abs(float(insertion_point[1])), 0)  # (x, y, z)
        # circle = space.AddCircle(center, cloud_radius)

        if HasAttributes:
            for attrib in entity.GetAttributes():
                if "DWGNO" in attrib.TagString:
                    print(attrib.TextString.strip())

                if "TAG" in attrib.TagString:
                    print(attrib.TextString.strip())

                if "DRAWING" in attrib.TagString and attrib.TextString:
                    print(attrib.TextString.strip())


            # Getting scale of block
            scale_x = entity.XScaleFactor
            scale_y = entity.YScaleFactor
            scale_z = entity.ZScaleFactor

            draw_circle(coord_list, (0.2 * float(scale_x)), color=4)




