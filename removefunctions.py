import zipfile
import re
import os
import time
from tempfile import mkdtemp
from shutil import rmtree, copyfile

def simulate_step(callback, message, step_count, total_steps):
    """Simulate a step in the process with a time delay."""
    time.sleep(1)  # Simulate time delay for each step
    progress = 100 * (step_count / total_steps)
    callback(message, progress)

def remove_workbook_protection(file_path, callback):
    total_steps = 17  # Adjust the total steps based on your process
    step_count = 0

    try:
        step_count += 1
        simulate_step(callback, "Creating temporary directory...", step_count, total_steps)
        temp_dir = mkdtemp()

        step_count += 1
        simulate_step(callback, "Creating backup for safety...", step_count, total_steps)
        backup_file_path = file_path + '.backup'
        copyfile(file_path, backup_file_path)

        step_count += 1
        simulate_step(callback, f"Extracting file at {temp_dir}...", step_count, total_steps)
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        step_count += 1
        simulate_step(callback, f"Reading xml content...", step_count, total_steps)
        with open(workbook_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        step_count += 1
        simulate_step(callback, "Removing workbook protection...", step_count, total_steps)
        xml_content = re.sub(r'<workbookProtection[^>]*>', '', xml_content)

        step_count += 1
        simulate_step(callback, "Saving modifications...", step_count, total_steps)
        with open(workbook_path, 'w', encoding='utf-8') as file:
            file.write(xml_content)

        step_count += 1
        simulate_step(callback, "Zipping file again...", step_count, total_steps)
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)

        rmtree(temp_dir)
        step_count += 1
        simulate_step(callback, "Workbook protection successfully removed.", step_count, total_steps)

        step_count += 1
        simulate_step(callback, "Creating temporary directory...", step_count, total_steps)
        temp_dir = mkdtemp()

        step_count += 1
        simulate_step(callback, f"Extracting file at {temp_dir}...", step_count, total_steps)
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        sheet_dir = os.path.join(temp_dir, "xl", "worksheets")
        for sheet_file in os.listdir(sheet_dir):
            if sheet_file.endswith('.xml'):
                sheet_path = os.path.join(sheet_dir, sheet_file)

                step_count += 1
                simulate_step(callback, f"Reading xml content...", step_count, total_steps)
                with open(sheet_path, 'r', encoding='utf-8') as file:
                    xml_content = file.read()

                step_count += 1
                simulate_step(callback, "Removing sheet protection...", step_count, total_steps)
                xml_content = re.sub(r'<sheetProtection[^>]*>', '', xml_content)

                step_count += 1
                simulate_step(callback, "Saving modifications...", step_count, total_steps)
                step_count += 1
                simulate_step(callback, "Zipping file again...", step_count, total_steps)
                step_count += 1
                simulate_step(callback, "Sheet protection successfully removed from all worksheets.", step_count, total_steps)
                step_count += 1
                with open(sheet_path, 'w', encoding='utf-8') as file:
                    file.write(xml_content)

        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, temp_dir)
                    zip_out.write(full_path, arcname)

        rmtree(temp_dir)

    except Exception as e:
        callback(f"An error occurred: {e}")
