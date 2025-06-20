import os
import subprocess
import shutil

input_folder = os.path.dirname(__file__)
output_folder = os.path.join(input_folder, "converted_docx")
os.makedirs(output_folder, exist_ok=True)

libreoffice_path = shutil.which("libreoffice")
if not libreoffice_path:
    # Try the default Mac path
    mac_soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac_soffice):
        libreoffice_path = mac_soffice
    else:
        print("LibreOffice is not installed or not found in PATH.")
        exit(1)

for filename in os.listdir(input_folder):
    if filename.lower().endswith(".doc"):
        input_path = os.path.join(input_folder, filename)
        result = subprocess.run([
            libreoffice_path,
            "--headless",
            "--convert-to", "docx",
            input_path,
            "--outdir", output_folder
        ], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"Converted {filename} to DOCX.")
        else:
            print(f"Failed to convert {filename}: {result.stderr}")