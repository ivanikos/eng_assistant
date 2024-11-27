import win32com.client
from pyautocad import Autocad, APoint




def draw_marker(center, radius, color=1):
    acad_1 = Autocad(create_if_not_exists=True)
    # Access the active document
    doc_1 = acad_1.ActiveDocument

    center_point = APoint(center[0], center[1], center[2])
    circle = doc_1.PaperSpace.AddCircle(center_point, radius)
    circle.Color = color
    return





acad = win32com.client.Dispatch("AutoCAD.Application")
acad.Visible

doc = acad.ActiveDocument  # Document object
doc_filename = doc.FullName

borders = {}


for entity in acad.ActiveDocument.PaperSpace:
    name = entity.EntityName

    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        if HasAttributes:
            load_tag = "n/a"
            for attrib in entity.GetAttributes():
                # print(f"HANDLE {entity.Handle}, TAG {attrib.TagString.strip()}, TEXT {attrib.TextString.strip()}")

                if attrib.TagString.strip() == "DWGNO":
                    borders[entity.Handle] = [attrib.TextString.strip(), entity.XScaleFactor, entity.InsertionPoint]
del acad



for i in borders:
    print(i, borders[i])

    border_scale = borders[i][1]
    border_dwg_number = borders[i][0]
    border_ins_point = borders[i][2]

    border_coord_list = [float(border_ins_point[0]), float(border_ins_point[1]), float(border_ins_point[2])]
    border_coord_list_x2 = [(float(border_ins_point[0]) + (42 * border_scale)), float(border_ins_point[1]), float(border_ins_point[2])]
    border_coord_list_y2 = [float(border_ins_point[0]), (float(border_ins_point[1]) + (30 * border_scale)), float(border_ins_point[2])]
    draw_marker(border_coord_list, 0.5 * float(border_scale), color=4)
    draw_marker(border_coord_list_x2, 0.5 * float(border_scale), color=4)
    draw_marker(border_coord_list_y2, 0.5 * float(border_scale), color=4)

# scale_x = target_entity.XScaleFactor
# scale_y = target_entity.YScaleFactor
# scale_z = target_entity.ZScaleFactor
# insertion_point = target_entity.InsertionPoint
#
# coord_list = [float(insertion_point[0]), float(insertion_point[1]), float(insertion_point[2])]
#
# draw_marker(coord_list, 0.2 * float(scale_x), color=4)