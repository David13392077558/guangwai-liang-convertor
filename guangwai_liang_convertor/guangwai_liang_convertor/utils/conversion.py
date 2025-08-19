import platform
import subprocess
import os

def convert_excel_to_pdf(input_path, output_path):
    system = platform.system()
    if system == "Windows":
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(input_path)
        wb.ExportAsFixedFormat(0, output_path)
        wb.Close(False)
        excel.Quit()
    else:
        cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            input_path,
            "--outdir", os.path.dirname(output_path)
        ]
        subprocess.run(cmd, check=True)

