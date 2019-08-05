import os, re
import win32com.client as win32
from win32com.client import constants
def save_as_xlsx(path):
    # Opening Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    doc = excel.Workbooks.Open(path)
    doc.Activate()

    # Rename path with .xlsx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub('\.\w+$', '.xlsx', new_file_abs)

    # Save and Close
    doc.SaveAs(
        new_file_abs, FileFormat=constants.xlOpenXMLStrictWorkbook
    )
    doc.Close(False)