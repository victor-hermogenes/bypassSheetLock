import zipfile
import re
from tempfile import mkdtemp
from shutil import rmtree, copyfile
import os

def remove_workbook_protection(file_path):
    try:
        temp_dir = mkdtemp()

        # Create a backup of the original file
        backup_file_path = file_path + '.backup'
        copyfile(file_path, backup_file_path)
        print(f"Backup created at: {backup_file_path}")

        # Open the Excel file as a zip file and extract its contents
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        
        # Read the workbook.xml file
        with open(workbook_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        # Use regex to find and remove the workbookProtection element
        xml_content = re.sub(r'<workbookProtection[^>]*>', '', xml_content)

        # Write the modified content back to workbook.xml
        with open(workbook_path, 'w', encoding='utf-8') as file:
            file.write(xml_content)

        # Re-zip the extracted files back into the .xlsx file
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)

        rmtree(temp_dir)
        print("Workbook protection successfully removed.")

    except Exception as e:
        print(f"An error occurred: {e}")


def remove_sheet_protection(file_path):
    try:
        temp_dir = mkdtemp()

        # Open the Excel file as a zip file and extract its contents
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        sheet_dir = os.path.join(temp_dir, "xl", "worksheets")
        for sheet_file in os.listdir(sheet_dir):
            if sheet_file.endswith('.xml'):
                sheet_path = os.path.join(sheet_dir, sheet_file)

                # Read the sheet XML file
                with open(sheet_path, 'r', encoding='utf-8') as file:
                    xml_content = file.read()

                # Use regex to find and remove the sheetProtection element
                xml_content = re.sub(r'<sheetProtection[^>]*>', '', xml_content)

                # Write the modified content back to the sheet XML file
                with open(sheet_path, 'w', encoding='utf-8') as file:
                    file.write(xml_content)

        # Re-zip the extracted files back into the .xlsx file
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)

        rmtree(temp_dir)
        print("Sheet protection successfully removed from all worksheets.")

    except Exception as e:
        print(f"An error occurred: {e}")