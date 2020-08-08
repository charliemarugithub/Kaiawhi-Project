import openpyxl as xl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import Workbook

# creating a variable to take any filename input from user
filename = input('Enter filename here: ')
new_sheet_name = input('Enter new sheet name: ')

# loading workbook on local computer c drive using filename
wb = xl.load_workbook(f'c:\\Users\\Charlie\\Desktop\\{filename}.xlsx')

# working with sheet1 on wb 'workbook'
sheet = wb['Form responses 3']

new_sheet = wb.create_sheet("Sheet A", 0)
new_sheet.title = new_sheet_name
new_sheet_name = wb.active


# defining packing list function
def packing_list():
    # deleting columns so that columns required are left for new file
    sheet.delete_cols(1, 9)
    sheet.delete_cols(5)

    # updating Column Names
    sheet['E1'].value = "Total"
    sheet['F1'].value = "Children"
    sheet['G1'].value = "Adults"
    sheet['I1'].value = "Are there any items you dont want included?"

    # calculate total number of rows and columns in source excel file
    max_rows = sheet.max_row
    max_columns = sheet.max_column

    # setting variables for loop
    bold_font = Font(name='Arial', size=12, bold=True)
    # cell alignment to center
    horizon_center = Alignment(horizontal='center')
    # row 1 align cells to vertical center
    vertical_center = Alignment(vertical='center')
    wrap_text = Alignment(wrap_text=True)
    # setting color for Panmure
    col_panmure = PatternFill(fgColor='bd291e', bgColor='FFFFFF', fill_type='solid')

    # copying the cell values from source excel file to destination excel file
    for i in range(1, max_rows + 1):
        for j in range(1, max_columns + 1):
            # reading cell value from source excel file
            c = sheet.cell(row=i, column=j)

            # writing the read value to destination excel file
            new_sheet.cell(row=i, column=j).value = c.value

            if new_sheet.cell(row=i, column=j).value == 'Panmure':
                print(i, j)
                new_sheet[i][j].fill = col_panmure

            # making row 1 bold font
            new_sheet.cell(row=1, column=i).font = bold_font
            # wrapping text on columns 8-10
            new_sheet.cell(row=1, column=8).alignment = wrap_text
            new_sheet.cell(row=i, column=9).alignment = wrap_text
            new_sheet.cell(row=i, column=10).alignment = wrap_text
            # text alignment for row 1
            new_sheet.cell(row=1, column=j).alignment = vertical_center
            new_sheet.cell(row=1, column=j).alignment = horizon_center

    # setting each column width to size
    # set the width of the column
    # I need to learn to iterate over this properly and set auto_size
    new_sheet_name.column_dimensions['A'].width = 21.5
    new_sheet_name.column_dimensions['B'].width = 35
    new_sheet_name.column_dimensions['C'].width = 27
    new_sheet_name.column_dimensions['D'].width = 25
    new_sheet_name.column_dimensions['E'].width = 9
    new_sheet_name.column_dimensions['F'].width = 12.5
    new_sheet_name.column_dimensions['G'].width = 9.8
    new_sheet_name.column_dimensions['H'].width = 85
    new_sheet_name.column_dimensions['I'].width = 85
    new_sheet_name.column_dimensions['J'].width = 120

    # saving new worksheet to desktop with name packing_list
    wb.remove_sheet(sheet)
    wb.save('c:\\Users\\Charlie\\Desktop\\packing_list.xlsx')
