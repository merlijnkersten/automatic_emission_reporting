from openpyxl import load_workbook
#from column_row_data import blablah

excel_path = r"C:/Users/Merlijn Kersten/Documents/Univerzita Karlova/Automatic emission reporting/Emission reporting template 1.xlsx"
workbook = load_workbook(filename=excel_path)

sheet_name = "ANNEX IV A-WM"

try:
    sheet = workbook[sheet_name]
except KeyError:
    print(f"No sheet named '{sheet_name}' in workbook")

'''
Next steps
* Import lists/dicts from column_row_data
* Import data from VEDA files
    * Need to cleanly import data
* Parse data from VEDA files to Excel file 
    * Need translation from VEDA/TIMES variable names to Excel names
'''
