from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def match_people():
    Tk().withdraw()
    fn = askopenfilename(title="Select an Excel file", filetypes=(("Excel Files", "*.xlsx*"), ("All Files", "*.*")))

    print(fn)
    if fn == '':
        print("No file selected. Exiting...")
        return

    wb = load_workbook(fn, read_only=True)
    sheet = wb.worksheets[0]
    print("Calculating worksheet dimensions...")
    sheet_dimensions = sheet.calculate_dimension().split(':')
    if sheet_dimensions[0] != 'A1' or sheet_dimensions[1] == 'A1':
        print("Warning: worksheet dimensions appear to be incorrect, resetting dimensions and recalculating...")
        sheet.reset_dimensions()
        try:
            sheet.calculate_dimension(force=True)
        except UnboundLocalError:
            print("It appears you may be using an empty (or very broken) Excel spreadsheet."
                  "Try adding some data or using one that's not broken. Exiting...")
            return

    print("Dimensions calculated successfully. Beginning pairing process...")
    for row in sheet.rows:
        for cell in row:
            print(cell.value)
    wb.close()


if __name__ == '__main__':
    match_people()
