import os
import zipfile
import xml.etree.ElementTree as ET

# Define the namespaces used in the workbook.xml file
namespaces = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def get_sheet_filename(workbook_xml, sheet_name):
    # Parse the workbook.xml file
    tree = ET.parse(workbook_xml)
    root = tree.getroot()

    # Find the sheet element by name
    sheet_elem = root.find(f"main:sheets/main:sheet[@name='{sheet_name}']", namespaces=namespaces)
    if sheet_elem is None:
        raise ValueError(f"Sheet named '{sheet_name}' not found in the workbook.")

    # Get the r:id attribute of the sheet element
    sheet_rid = sheet_elem.attrib[f"{{{namespaces['r']}}}id"]

    return sheet_rid

# Assuming the workbook.xml has been extracted to 'extracted_folder'
workbook_xml_path = os.path.join('extracted_folder', 'xl', 'workbook.xml')
sheet_name = 'Base CEP Km + Int e cap + Prazo'  # Example sheet name
sheet_rid = get_sheet_filename(workbook_xml_path, sheet_name)

# Output the r:id
print(f"The r:id for the sheet '{sheet_name}' is {sheet_rid}")