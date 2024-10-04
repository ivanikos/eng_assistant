# from pyautocad import Autocad, APoint
import win32com.client
import math
from array import array



acad = win32com.client.Dispatch("AutoCAD.Application")
print(acad.FullName)

print(acad)

doc = acad.ActiveDocument  # Document object
space = doc.PaperSpace
print("Paperspace - ",len(acad.ActiveDocument.PaperSpace))
print("Blocks - ", len(doc.Blocks))

# Parameters for cloud (number of bumps and size)
num_bumps = 8  # Number of arcs to form the cloud
cloud_radius = 10  # Radius of the cloud from the insertion point


# Function to create a circular cloud shape around a given point
def draw_cloud(center, radius, num_bumps):
    angle_step = 2 * math.pi / num_bumps
    points = []

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
    polyline = space.AddPolyline(polyline_points)
    polyline.Closed = True  # Close the polyline to make it a loop


for entity in acad.ActiveDocument.PaperSpace:
    name = entity.EntityName

    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        insertion_point = entity.InsertionPoint

        center = (abs(float(insertion_point[0])), abs(float(insertion_point[1])), 0)  # (x, y, z)
        circle = space.AddCircle(center, cloud_radius)

        if HasAttributes:
            for attrib in entity.GetAttributes():
                if "DWGNO" in attrib.TagString:
                    print(attrib.TextString.strip())

                if "TAG" in attrib.TagString:
                    print(attrib.TextString.strip())

                if "DRAWING" in attrib.TagString and attrib.TextString:
                    print(attrib.TextString.strip())






            # draw_cloud(entity.InsertionPoint, 5, 10)




"""
from pyautocad import Autocad, APoint

# Create an instance of AutoCAD
acad = Autocad(create_if_not_exists=True)

# Parameters for the cloud circle
cloud_radius = 10  # Radius of the cloud circle

# Iterate through all objects in the active document
for entity in acad.iter_objects("AcDbBlockReference"):
    # Get the insertion point of the block
    insertion_point = entity.InsertionPoint
    center = APoint(insertion_point)  # Convert to APoint

    # Draw a circle around the block's insertion point
    circle = acad.model.AddCircle(center, cloud_radius)
    print(f"Circle drawn around block '{entity.Name}' at {insertion_point}")


"""
