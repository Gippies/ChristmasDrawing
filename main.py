from openpyxl import load_workbook


def match_people():
    wb = load_workbook("TestFile.xlsx", read_only=True)
    sheet = wb.worksheets[0]
    print("Calculating worksheet dimensions...")
    sheet_dimensions = sheet.calculate_dimension().split(':')
    if sheet_dimensions[0] != 'A1' or sheet_dimensions[1] == 'A1':
        print("Warning: worksheet dimensions appear to be incorrect, resetting dimensions and recalculating...")
        sheet.reset_dimensions()
        sheet.calculate_dimension(force=True)
    for row in sheet.rows:
        for cell in row:
            print(cell.value)
    wb.close()


if __name__ == '__main__':
    match_people()
