import os
import tkinter
import zipfile
import xml.etree.ElementTree as ET


import tkinter.messagebox
from modulefinder import packagePathMap

import customtkinter
import tkinterdnd2
import threading
from tkinter.filedialog import askopenfile, asksaveasfilename

from modules.sm_logic import get_pcat_list, change_spec_paths

# Functions section-------------------------------------------------------------

# specs_folder = r"C:\Users\ivaign\DC\ACCDocs\UCC\00000-00\Project Files\13 - Plant 3D\00000-00\Spec Sheets\\"
specs_folder = r"C:\Users\ivaign\DC\ACCDocs\UCC\UCC Template\Project Files\Plant 3D\_PLANT 3D 2025 IMPERIAL TEMPLATE\_PLANT 3D 2025 IMPERIAL TEMPLATE\Spec Sheets\\"
content_folder = r"C:\Users\ivaign\DC\ACCDocs\UCC\UCC Template\Project Files\Plant 3D\AutoCAD Plant 3D 2025 Content"


pcat_paths ={}
original_pcat_paths_list = {}
not_ok = []


# Functions section end -------------------------------------------------------------
def migrate(spec_path, pcat_paths):
    change_spec_paths(spec_path, pcat_paths)


for file_name in os.listdir(specs_folder):
    pcat_paths = {}
    if ".pspx" in file_name:
        print(specs_folder + file_name)
        print((specs_folder + file_name).split("\\")[-1])

        with zipfile.ZipFile((specs_folder + file_name), 'r') as zip_ref:
            # Iterate through all files in the original zip
            for item in zip_ref.infolist():
                # If the current file is CatalogReferences.xml, modify it
                if item.filename == 'editor/CatalogReferences.xml':
                    # Extract the XML file into memory
                    with zip_ref.open(item.filename) as xml_file:
                        tree = ET.parse(xml_file)
                        root = tree.getroot()
                        for elem in root.iter():
                            # Make modifications to the XML (example: modify tag)
                            if "pcat" in elem.text or "acat" in elem.text:
                                pcat_name = elem.text.split("\\")[-1]
                                old_pcat_path = elem.text.replace(pcat_name, "")
                                # print(elem.text.replace("N:\Piping\AutoCAD Plant 3D 2025 Content", content_folder))
                                print(elem.text)

                                # elem.text = elem.text.replace(old_pcat_path,
                                #                               pcat_paths[pcat_name][0].replace("/", "\\"))



        catalog_name = ''


