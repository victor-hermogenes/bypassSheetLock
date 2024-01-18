import zipfile
import os
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
from shutil import rmtree, copyfile

file_path = input('Caminho para a planilha: ')

def bypassSheetLock(file_path):
    try:
        print(f"Reading Excel file {file_path}")
        namespaces = {
            '', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            # add additional namespace decalration if needed
        }
        namespace_uri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        temp_dir = mkdtemp()
        modified_files = []

        # Extract files
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Find workbookProtection
        print("Looking for workbook Protection...")
        workbook_path = f"{temp_dir}/xl/workbook.xml"
        workbook_tree = ET.parse(workbook_path)
        workbook_root = workbook_tree.getroot()
        workbook_protection = workbook_root.find(f'{{{namespace_uri}}}workbookProtection')

        # Delete workbookProtection
        if workbook_protection is not None:
                print("Removing workbook protection...")
                workbook_root.remove(workbook_protection)
                workbook_tree.write(workbook_path)
                modified_files.append(workbook_path)
                print("Worbook protection deleted.")

        # Find sheetProtection
        print("Looking for worksheet protection.")
        sheet_dir = f"{temp_dir}/xl/worksheets"
        for sheet_file in os.listdir(sheet_dir):
            if sheet_file.endswith('.xml'):
                sheet_path = f"{sheet_dir}/{sheet_file}"
                sheet_tree = ET.parse(sheet_path)
                sheet_root = sheet_tree.getroot()
                sheet_protection = sheet_root.find(f'{{{namespace_uri}}}sheetProtection')

                # Delete sheetProtection
                if sheet_protection is not None:
                    print("Removing worksheet protection...")
                    sheet_root.remove(sheet_protection)
                    sheet_tree.write(sheet_path)
                    modified_files.append(sheet_path)
                    print("Worksheet protection deleted.")

        # Create a backup of the original file for safe measures
        backup_file_path = file_path + '.backup'
        print(f"Creating backup at {backup_file_path} for safety measures...")
        copyfile(file_path, backup_file_path)
        print(f"Backup created at: {backup_file_path}.")

        # Re-zip files, only modifying the files that were changed
        print("Trying to rezip the file...")
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    if full_path in modified_files:
                        zip_out.write(full_path, arcname)
                    else:
                        with open(full_path, 'rb') as file_data:
                            zip_out.writestr(os.path.relpath(full_path, temp_dir), file_data.read())

        # Warning
        print("File rezipped.")
        rmtree(temp_dir)
        print("Protection removal process completed!")
    except Exception as e:
        print(f"Error: {e}")

# Test
bypassSheetLock(file_path)