import zipfile
import os
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
from shutil import rmtree, copyfile

def bypassSheetLock(file_path):
    try:
        print(f"Reading Excel file {file_path}")

        # Namespace declarations
        namespaces = {
            '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            # Add additional namespaces if needed
        }

        # Register namespaces
        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        temp_dir = mkdtemp()
        modified_files = []

        # Extract files
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Find workbookProtection
        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        workbook_tree = ET.parse(workbook_path)
        workbook_root = workbook_tree.getroot()

        # Find the workbookProtection element
        workbook_protection = workbook_root.find('workbookProtection', namespaces)
        if workbook_protection is not None:
            print("Removing workbook protection...")
            workbook_root.remove(workbook_protection)
            workbook_tree.write(workbook_path, encoding='UTF-8', xml_declaration=True)
            modified_files.append(workbook_path)
        else:
            print("workbookProtection not found...")

        # Find and remove sheetProtection
        sheet_dir = os.path.join(temp_dir, "xl", "worksheets")
        for sheet_file in os.listdir(sheet_dir):
            if sheet_file.endswith('.xml'):
                sheet_path = os.path.join(sheet_dir, sheet_file)
                sheet_tree = ET.parse(sheet_path)
                sheet_root = sheet_tree.getroot()
                sheet_protection = sheet_root.find('.//sheetProtection', namespaces)
                if sheet_protection is not None:
                    print(f"Removing protection from {sheet_file}...")
                    parent_element = sheet_protection.find('..')
                    if parent_element is not None:
                        parent_element.remove(sheet_protection)
                        sheet_tree.write(sheet_path, encoding='UTF-8', xml_declaration=True)
                        modified_files.append(sheet_path)
                else:
                    print("sheetProtection not found")

        # Create a backup of the original file
        backup_file_path = file_path + '.backup'
        copyfile(file_path, backup_file_path)
        print(f"Backup created at: {backup_file_path}")

        # Re-zip files
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname) if full_path in modified_files else zip_out.writestr(arcname, open(full_path, 'rb').read())

        # Clean up
        rmtree(temp_dir)
        print("Protection removal process completed.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Usage
file_path = "C:\\Users\\Master\\Desktop\\Tabelas importadoras\\Rodonaves\\TestPythonCode.xlsx"
bypassSheetLock(file_path)
