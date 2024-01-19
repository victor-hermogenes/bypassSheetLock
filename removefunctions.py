import zipfile
import re
import os
import time
from tempfile import mkdtemp
from shutil import rmtree, copyfile

total_steps = 30
step_count = 0

def remove_workbook_protection(file_path, callback):
    global step_count
    try:
        step_count += 1
        callback(100 * (step_count / total_steps),"Creating temporary directory...")
        temp_dir = mkdtemp()
        step_count += 1
        callback(100 * (step_count / total_steps),f"Directory {temp_dir} created!")

        # Create a backup of the original file
        step_count += 1
        callback(100 * (step_count / total_steps),"Creating backup for safety...")
        backup_file_path = file_path + '.backup'
        copyfile(file_path, backup_file_path)
        step_count += 1
        callback(100 * (step_count / total_steps),f"Backup created at: {backup_file_path}!")

        # Open the Excel file as a zip file and extract its contents
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        step_count += 1
        callback(100 * (step_count / total_steps),f"Extracting file at {temp_dir}...")

        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        
        # Read the workbook.xml file
        with open(workbook_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
        step_count += 1
        callback(100 * (step_count / total_steps),f"Reading xml content: {xml_content}...")
        step_count += 1
        callback(100 * (step_count / total_steps),f"Looking for workbook protection...")

        # Use regex to find and remove the workbookProtection element
        step_count += 1
        callback(100 * (step_count / total_steps),"Workbook protection found!")
        step_count += 1
        callback(100 * (step_count / total_steps),"Removing workbook protection...")
        xml_content = re.sub(r'<workbookProtection[^>]*>', '', xml_content)
        step_count += 1
        callback(100 * (step_count / total_steps),"Workbook protection removed!")


        # Write the modified content back to workbook.xml
        with open(workbook_path, 'w', encoding='utf-8') as file:
            step_count += 1
            callback(100 * (step_count / total_steps),"Saving modifications...")
            file.write(xml_content)
            step_count += 1
            callback(100 * (step_count / total_steps),"Modifications saved!")

        # Re-zip the extracted files back into the .xlsx file
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            step_count += 1
            callback(100 * (step_count / total_steps),"Zipping file again...")
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)
            step_count += 1
            callback(100 * (step_count / total_steps),"File zipped!")

        rmtree(temp_dir)
        step_count += 1
        callback(100 * (step_count / total_steps),"Workbook protection successfully removed.")

    except Exception as e:
        callback(f"An error occurred: {e}")


def remove_sheet_protection(file_path, callback):
    global step_count
    try:
        step_count += 1
        callback(100 * (step_count / total_steps),"Creating temporary directory...")
        temp_dir = mkdtemp()
        step_count += 1
        callback(100 * (step_count / total_steps),f"Directory {temp_dir} created!")

        # Open the Excel file as a zip file and extract its contents
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            step_count += 1
            callback(100 * (step_count / total_steps),f"Extracting file at {temp_dir}...")

        sheet_dir = os.path.join(temp_dir, "xl", "worksheets")
        
        for sheet_file in os.listdir(sheet_dir):
            if sheet_file.endswith('.xml'):
                sheet_path = os.path.join(sheet_dir, sheet_file)

                # Read the sheet XML file
                with open(sheet_path, 'r', encoding='utf-8') as file:
                    xml_content = file.read()
                step_count += 1
                callback(100 * (step_count / total_steps),f"Reading xml content: {xml_content}...")
                step_count += 1
                callback(100 * (step_count / total_steps),f"Looking for sheet protection...")
                
                # Use regex to find and remove the sheetProtection element
                step_count += 1
                callback(100 * (step_count / total_steps),"Sheet protection found!")
                step_count += 1
                callback(100 * (step_count / total_steps),"Removing sheet protection...")
                xml_content = re.sub(r'<sheetProtection[^>]*>', '', xml_content)
                step_count += 1
                callback(100 * (step_count / total_steps),"Sheet protection removed!")

                # Write the modified content back to the sheet XML file
                with open(sheet_path, 'w', encoding='utf-8') as file:
                    step_count += 1
                    callback(100 * (step_count / total_steps),"Saving modifications...")
                    file.write(xml_content)
                    step_count += 1
                    callback(100 * (step_count / total_steps),"Modifications saved!")

        # Re-zip the extracted files back into the .xlsx file
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            step_count += 1
            callback(100 * (step_count / total_steps),"Zipping file again...")
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)
            step_count += 1
            callback(100 * (step_count / total_steps),"File zipped!")

        rmtree(temp_dir)
        step_count += 1
        callback(100 * (step_count / total_steps),"Sheet protection successfully removed from all worksheets.")

    except Exception as e:
        callback(f"An error occurred: {e}")