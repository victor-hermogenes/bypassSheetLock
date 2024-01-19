import zipfile
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
from shutil import rmtree
import os

file_path = "C:\\Users\\Master\\Desktop\\Tabelas importadoras\\Rodonaves\\TestPythonCode.xlsx"

def find_workbook_protection(file_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/workbook.xml') as workbook_file:
                workbook_tree = ET.parse(workbook_file)
                workbook_root = workbook_tree.getroot()

                namespaces = {
                    '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                    # Add other namespaces if necessary
                }

                # Find the workbookProtection element
                workbook_protection = workbook_root.find('workbookProtection', namespaces)

                # Print details of the workbookProtection element
                if workbook_protection is not None:
                    print("workbookProtection element found:")
                    for attribute, value in workbook_protection.attrib.items():
                        print(f"  {attribute}: {value}")
                else:
                    print("workbookProtection element not found.")
                
                return workbook_protection

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def remove_workbook_protection(file_path):
    try:
        # Use find_workbook_protection to get the workbookProtection element
        workbook_protection = find_workbook_protection(file_path)
        if workbook_protection is None:
            print("workbookProtection element not found. No changes made.")
            return

        # Create a temporary directory
        temp_dir = mkdtemp()

        # Open the Excel file as a zip file and extract its contents
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Modify the workbook.xml file
        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        workbook_tree = ET.parse(workbook_path)
        workbook_root = workbook_tree.getroot()

        # Remove the workbookProtection element
        print("Removing workbook protection...")
        workbook_root.remove(workbook_protection)
        workbook_tree.write(workbook_path, encoding='UTF-8', xml_declaration=True)

        # Re-zip the extracted files back into the .xlsx file
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)

        # Clean up the temporary directory
        rmtree(temp_dir)
        print("Workbook protection successfully removed.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Usage example
remove_workbook_protection(file_path)
