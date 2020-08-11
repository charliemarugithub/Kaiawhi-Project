import openpyxl as xl
import tkinter as tk
from tkinter import ttk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.borders import BORDER_THICK
from functions import file_not_found, no_filename, no_sheet_name
from openpyxl.utils.exceptions import InvalidFileException

# creating instance of TK class
root = tk.Tk()
# creating main form window and packing it
main_form = tk.Canvas(root, width=500, height=400)
main_form.pack()
main_form.configure(background="#7289f2")
# name root title page
root.title('Kaiawhi Program')

# icon
root.iconbitmap('D:\Kaiawhi Folder\kaiawhiconImg.ico')
# creating label for filename and placing it in root
filename_label = tk.Label(root, text='Enter source file name: ', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 40, window=filename_label)

# creating entry for filename and placing it in root
filename_entry = tk.Entry(root, font="Helvetica, 16")
main_form.create_window(250, 80, window=filename_entry, width=220, height=25)

# creating label for sheet name and placing it in root
sheet_name_label = tk.Label(root, text='Enter new sheet name:', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 120, window=sheet_name_label)

# creating entry for sheet name  and placing it in root
sheet_name_example_label = tk.Label(root, text='eg 10 Aug Packing List', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 150, window=sheet_name_example_label)

# creating entry for sheet name and placing it in root
sheet_name_entry = tk.Entry(root, font="Helvetica, 16")
main_form.create_window(250, 185, window=sheet_name_entry, width=220, height=25)


def make_packing_list():
    # get method for filename entry
    x1 = filename_entry.get()

    try:
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

    except FileNotFoundError:
        file_not_found()

    except InvalidFileException:
        no_filename()

    except ValueError:
        no_sheet_name()


def make_delivery_list():
    # get method for filename entry
    x1 = filename_entry.get()
    try:
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
        sheet.delete_cols(1, 7)
        sheet.delete_cols(9, 4)

        # updating Column Names
        sheet['A1'].value = "First Name"
        sheet['C1'].value = "Address Number"
        sheet['G1'].value = "Delivery Instructions"
        sheet['H1'].value = "Total"

        # calculate total number of rows and columns in source excel file
        max_rows = sheet.max_row
        max_columns = sheet.max_column

        # setting variables for loop
        bold_font = Font(name='Arial', size=12, bold=True)
        # cell alignment to center
        horizon_center = Alignment(horizontal='center')
        # wrap text alignment
        wrap_text = Alignment(wrap_text=True)
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
                # wrapping text on columns 9
                new_sheet.cell(row=i, column=9).alignment = wrap_text
                # setting 'Totals' color column to red and bold font
                new_sheet.cell(row=i, column=8).fill = totals_color
                new_sheet.cell(row=i, column=8).font = bold_font
                # setting specific column widths
                x2.column_dimensions['A'].width = 18
                x2.column_dimensions['B'].width = 28
                x2.column_dimensions['C'].width = 35
                x2.column_dimensions['D'].width = 35
                x2.column_dimensions['E'].width = 30
                x2.column_dimensions['F'].width = 25
                x2.column_dimensions['G'].width = 60
                x2.column_dimensions['H'].width = 12
                x2.column_dimensions['I'].width = 75

        # saving new worksheet to desktop with name packing_list
        wb.remove(sheet)
        wb.save('c:\\Users\\Charlie\\Desktop\\delivery_list.xlsx')

    except FileNotFoundError:
        file_not_found()

    except InvalidFileException:
        no_filename()

    except ValueError:
        no_sheet_name()


packing_button = ttk.Button(text='Packing List', command=make_packing_list)
main_form.create_window(150, 260, window=packing_button, height=50, width=150)

delivery_button = ttk.Button(text='Delivery List', command=make_delivery_list)
main_form.create_window(350, 260, window=delivery_button, height=50, width=150)

root.mainloop()
