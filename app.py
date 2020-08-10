import openpyxl as xl
import tkinter as tk
from tkinter import ttk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.borders import BORDER_THICK


root = tk.Tk()

main_form = tk.Canvas(root, width=400, height=300)
main_form.pack()

root.title('Kaiawhi Program')

filename_label = tk.Label(root, text='Enter filename: ')
main_form.create_window(200, 20, window=filename_label)

filename_entry = tk.Entry(root)
main_form.create_window(200, 40, window=filename_entry)

sheet_name_label = tk.Label(root, text='Enter new sheet name:')
main_form.create_window(200, 80, window=sheet_name_label)

sheet_name_example_label = tk.Label(root, text='eg 10 Aug Packing List')
main_form.create_window(200, 100, window=sheet_name_example_label)

sheet_name_entry = tk.Entry(root)
main_form.create_window(200, 130, window=sheet_name_entry)


def make_packing_list():
    # get method for filename entry
    x1 = filename_entry.get()

    # loading workbook on local computer c drive using filename
    wb = xl.load_workbook(f'c:\\Users\\Charlie\\Desktop\\{x1}.xlsx')

    # working with sheet1 on wb 'workbook'
    sheet = wb['Form responses 3']

    # get method for new sheet name
    x2 = sheet_name_entry.get()
    new_sheet = wb.create_sheet("Sheet A", 0)
    new_sheet.title = x2
    x2 = wb.active

# deleting columns so that columns required are left for new file
    sheet.delete_cols(1, 9)
    sheet.delete_cols(5)

    # updating Column Names
    sheet['E1'].value = "Total"
    sheet['F1'].value = "Children"
    sheet['G1'].value = "Adults"
    sheet['H1'].value = "Packing Instructions"
    sheet['I1'].value = "Are there any items you dont want included?"

    # calculate total number of rows and columns in source excel file
    max_rows = sheet.max_row
    max_columns = sheet.max_column

    # setting variables for loop
    bold_font = Font(name='Arial', size=12, bold=True)
    # cell alignment to center
    horizon_center = Alignment(horizontal='center')
    # wrap text alignment
    wrap_text = Alignment(wrap_text=True)
    # setting border types
    border = Border(
                left=Side(border_style=BORDER_THICK, color='a8a1ad'),
                right=Side(border_style=BORDER_THICK, color='a8a1ad'),
                top=Side(border_style=BORDER_THICK, color='a8a1ad'),
                bottom=Side(border_style=BORDER_THICK, color='a8a1ad')
                )
    # setting colors for each suburb
    col_panmure = PatternFill(fgColor='80e098', fill_type='solid')
    col_ptengland = PatternFill(fgColor='d9b36c', fill_type='solid')
    col_gi = PatternFill(fgColor='8d9cf0', fill_type='solid')
    col_stjohns = PatternFill(fgColor='ba6cd9', fill_type='solid')
    col_glendowie = PatternFill(fgColor='d98aed', fill_type='solid')
    col_mtwell = PatternFill(fgColor='a1f0a9', fill_type='solid')
    col_greenlane = PatternFill(fgColor='f0daa1', fill_type='solid')
    col_mangere = PatternFill(fgColor='e4f0a1', fill_type='solid')
    col_pakuranga = PatternFill(fgColor='e4e86d', fill_type='solid')
    # setting color for 'Total' column
    totals_color = PatternFill(fgColor='fac2b4', fill_type='solid')

    # copying the cell values from source excel file to destination excel file
    for i in range(1, max_rows + 1):
        for j in range(1, max_columns + 1):
            # reading cell value from source excel file
            c = sheet.cell(row=i, column=j)

            # writing the read value to destination excel file
            new_sheet.cell(row=i, column=j).value = c.value

            for row in x2['A1:J100']:
                for cell in row:
                    if cell.value == 'Panmure':
                        cell.fill = col_panmure
                    if cell.value == 'Pt England':
                        cell.fill = col_ptengland
                    if cell.value == 'Point England':
                        cell.fill = col_ptengland
                    if cell.value == 'Glen Innes':
                        cell.fill = col_gi
                    if cell.value == 'St Johns':
                        cell.fill = col_stjohns
                    if cell.value == 'Glendowie':
                        cell.fill = col_glendowie
                    if cell.value == 'Mt Wellington':
                        cell.fill = col_mtwell
                    if cell.value == 'Greenlane':
                        cell.fill = col_greenlane
                    if cell.value == 'Mangere':
                        cell.fill = col_mangere
                    if cell.value == 'Pakuranga':
                        cell.fill = col_pakuranga

            # making row 1 bold font
            new_sheet.cell(row=1, column=i).font = bold_font
            # text alignment for all rows
            new_sheet.cell(row=i, column=j).alignment = horizon_center
            # setting all row height to 30
            x2.row_dimensions[i].height = 30
            # setting borders for all cells
            new_sheet.cell(row=i, column=j).border = border
            # wrapping text on columns 8-10
            new_sheet.cell(row=i, column=8).alignment = wrap_text
            new_sheet.cell(row=i, column=9).alignment = wrap_text
            new_sheet.cell(row=i, column=10).alignment = wrap_text
            # setting 'Totals' color column to red and bold font
            new_sheet.cell(row=i, column=5).fill = totals_color
            new_sheet.cell(row=i, column=5).font = bold_font
            # setting specific column widths
            x2.column_dimensions['A'].width = 21.5
            x2.column_dimensions['B'].width = 35
            x2.column_dimensions['C'].width = 27
            x2.column_dimensions['D'].width = 25
            x2.column_dimensions['E'].width = 9
            x2.column_dimensions['F'].width = 12.5
            x2.column_dimensions['G'].width = 9.8
            x2.column_dimensions['H'].width = 75
            x2.column_dimensions['I'].width = 75
            x2.column_dimensions['J'].width = 75

    # saving new worksheet to desktop with name packing_list
    wb.remove(sheet)
    wb.save('c:\\Users\\Charlie\\Desktop\\packing_list.xlsx')


packing_button = ttk.Button(text='Packing List', command=make_packing_list)
main_form.create_window(150, 180, window=packing_button)

# delivery_button = ttk.Button(text='Delivery List', command=make_delivery_list)
# main_form.create_window(250, 180, window=delivery_button)

root.mainloop()
