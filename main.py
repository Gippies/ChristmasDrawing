from colorama import init as colorama_init, Fore
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def match_people():
    Tk().withdraw()
    fn = askopenfilename(title="Select an Excel file", filetypes=(("Excel Files", "*.xlsx*"), ("All Files", "*.*")))

    if fn == '':
        print(Fore.RED + "No file selected. Exiting...")
        return

    wb = load_workbook(fn, read_only=True)
    sheet = wb.worksheets[0]
    print("Calculating worksheet dimensions...")
    sheet_dimensions = sheet.calculate_dimension().split(':')
    if sheet_dimensions[0] != 'A1' or sheet_dimensions[1] == 'A1':
        print(Fore.YELLOW + "Warning: worksheet dimensions appear to be incorrect,"
                            "resetting dimensions and recalculating...")
        sheet.reset_dimensions()
        try:
            sheet.calculate_dimension(force=True)
        except UnboundLocalError:
            print(Fore.RED + "Error: It appears you may be using an empty (or very broken) Excel spreadsheet."
                  "Try adding some data or using one that's not broken. Exiting...")
            return

    print(Fore.GREEN + "Dimensions calculated successfully. Beginning pairing process...")
    people_dict = {}
    for row in sheet.rows:
        people_dict[row[0].value] = []
        for cell in row[1:]:
            if cell.value is not None:
                people_dict[row[0].value].append(cell.value)

    individuals = list(people_dict.keys())
    if len(individuals) > len(set(individuals)):
        print(Fore.RED + "Error: There are duplicate individuals in the 'A' column."
                         "If you have multiple exclusions, please put them all on the same row. Exiting...")
        return
    print("People to pair: " + str(individuals))
    print("People with exclusions: " + str(people_dict))
    wb.close()


if __name__ == '__main__':
    colorama_init(autoreset=True)
    match_people()
