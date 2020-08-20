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

# creating label for filename and placing it in root
filename_label = tk.Label(root, text='Enter source file path name: ', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 40, window=filename_label)

# creating entry for filename and placing it in root
filename_entry = tk.Entry(root, font="Helvetica, 16")
main_form.create_window(250, 80, window=filename_entry, width=350, height=25)

# creating label for sheet name and placing it in root
sheet_name_label = tk.Label(root, text='Enter new sheet name:', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 120, window=sheet_name_label)

# creating entry for sheet name  and placing it in root
sheet_name_example_label = tk.Label(root, text='eg 10 Aug Packing List', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 150, window=sheet_name_example_label)

# creating entry for sheet name and placing it in root
sheet_name_entry = tk.Entry(root, font="Helvetica, 16")
main_form.create_window(250, 185, window=sheet_name_entry, width=300, height=25)


def make_packing_list():
    # get method for filename entry
    get_file = filename_entry.get()

    try:
        # loading workbook on local computer c drive using filename
        wb = xl.load_workbook(f'{get_file.strip()}.xlsx')

        # working with sheet1 on wb 'workbook'
        sheet = wb['Form responses 3']

        # get method for new sheet name
        sheet_name = sheet_name_entry.get()
        new_sheet = wb.create_sheet("Sheet A", 0)
        new_sheet.title = sheet_name
        sheet_name = wb.active

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

        # setting font for row 1
        bold_font = Font(name='Arial', size=14, bold=True)
        # setting font for text in whole sheet except top row
        cell_font = Font(name='Calibri', size=16)
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
        col_clendon = PatternFill(fgColor='edc0b4', fill_type='solid')
        col_henderson = PatternFill(fgColor='eddcb4', fill_type='solid')
        col_howick = PatternFill(fgColor='dbedb4', fill_type='solid')
        col_karaka = PatternFill(fgColor='84b07f', fill_type='solid')
        col_manukau = PatternFill(fgColor='7fb099', fill_type='solid')
        col_manurewa = PatternFill(fgColor='98edc5', fill_type='solid')
        col_meadowbank = PatternFill(fgColor='77d1c8', fill_type='solid')
        col_onehunga = PatternFill(fgColor='87d0ed', fill_type='solid')
        col_otahuhu = PatternFill(fgColor='bebdf2', fill_type='solid')
        col_waiotaiki = PatternFill(fgColor='9292a6', fill_type='solid')
        col_wattle = PatternFill(fgColor='cebad6', fill_type='solid')
        # setting color for 'Total' column
        # totals_color = PatternFill(fgColor='fac2b4', fill_type='solid')

        # counter for number of times the suburb is selected
        counter_panmure = 0
        # copying the cell values from source excel file to destination excel file
        for i in range(1, max_rows + 1):
            for j in range(1, max_columns + 1):
                # reading cell value from source excel file
                c = sheet.cell(row=i, column=j)

                # writing the read value to destination excel file
                new_sheet.cell(row=i, column=j).value = c.value
                new_sheet.cell(row=i, column=j).font = cell_font

                for row in sheet_name.iter_rows(max_row=max_rows, max_col=max_columns):
                    for cell in row:
                        if cell.value == 'Panmure':
                            cell.fill = col_panmure
                            new_sheet.cell(row=i, column=5).fill = col_panmure
                        if cell.value == 'Clendon Park':
                            cell.fill = col_clendon
                            new_sheet.cell(row=i, column=5).fill = col_clendon
                        if cell.value == 'Point England':
                            cell.fill = col_ptengland
                            new_sheet.cell(row=i, column=5).fill = col_ptengland
                        if cell.value == 'Glen Innes':
                            cell.fill = col_gi
                            new_sheet.cell(row=i, column=5).fill = col_gi
                        if cell.value == 'St Johns':
                            cell.fill = col_stjohns
                            new_sheet.cell(row=i, column=5).fill = col_stjohns
                        if cell.value == 'Glendowie':
                            cell.fill = col_glendowie
                            new_sheet.cell(row=i, column=5).fill = col_glendowie
                        if cell.value == 'Mt Wellington':
                            cell.fill = col_mtwell
                            new_sheet.cell(row=i, column=5).fill = col_mtwell
                        if cell.value == 'Greenlane':
                            cell.fill = col_greenlane
                            new_sheet.cell(row=i, column=5).fill = col_greenlane
                        if cell.value == 'Mangere':
                            cell.fill = col_mangere
                            new_sheet.cell(row=i, column=5).fill = col_mangere
                        if cell.value == 'Pakuranga':
                            cell.fill = col_pakuranga
                            new_sheet.cell(row=i, column=5).fill = col_pakuranga
                        if cell.value == 'Henderson':
                            cell.fill = col_henderson
                            new_sheet.cell(row=i, column=5).fill = col_henderson
                        if cell.value == 'Howick':
                            cell.fill = col_howick
                            new_sheet.cell(row=i, column=5).fill = col_howick
                        if cell.value == 'Karaka':
                            cell.fill = col_karaka
                            new_sheet.cell(row=i, column=5).fill = col_karaka
                        if cell.value == 'Manukau':
                            cell.fill = col_manukau
                            new_sheet.cell(row=i, column=5).fill = col_manukau
                        if cell.value == 'Manurewa':
                            cell.fill = col_manurewa
                            new_sheet.cell(row=i, column=5).fill = col_manurewa
                        if cell.value == 'Meadowbank':
                            cell.fill = col_meadowbank
                            new_sheet.cell(row=i, column=5).fill = col_meadowbank
                        if cell.value == 'Onehunga':
                            cell.fill = col_meadowbank
                            new_sheet.cell(row=i, column=5).fill = col_onehunga
                        if cell.value == 'Otahuhu':
                            cell.fill = col_otahuhu
                            new_sheet.cell(row=i, column=5).fill = col_otahuhu
                        if cell.value == 'Waiotaiki Bay':
                            cell.fill = col_waiotaiki
                            new_sheet.cell(row=i, column=5).fill = col_waiotaiki
                        if cell.value == 'Wattle Downs':
                            cell.fill = col_wattle
                            new_sheet.cell(row=i, column=5).fill = col_wattle

                # making row 1 bold font
                new_sheet.cell(row=1, column=i).font = bold_font
                # text alignment for all rows
                new_sheet.cell(row=i, column=j).alignment = horizon_center
                # setting all row height to 30
                sheet_name.row_dimensions[i].height = 40
                # setting borders for all cells
                new_sheet.cell(row=i, column=j).border = border
                # wrapping text on columns 8-10
                new_sheet.cell(row=i, column=8).alignment = wrap_text
                new_sheet.cell(row=i, column=9).alignment = wrap_text
                new_sheet.cell(row=i, column=10).alignment = wrap_text
                # setting 'Totals' color column to red and bold font
                # new_sheet.cell(row=i, column=5).fill = totals_color
                new_sheet.cell(row=i, column=5).font = bold_font
                # setting specific column widths
                sheet_name.column_dimensions['A'].width = 20
                sheet_name.column_dimensions['B'].width = 30
                sheet_name.column_dimensions['C'].width = 25
                sheet_name.column_dimensions['D'].width = 25
                sheet_name.column_dimensions['E'].width = 9
                sheet_name.column_dimensions['F'].width = 13
                sheet_name.column_dimensions['G'].width = 10
                sheet_name.column_dimensions['H'].width = 45
                sheet_name.column_dimensions['I'].width = 45
                sheet_name.column_dimensions['J'].width = 45

        # saving new worksheet to desktop with name packing_list
        wb.remove(sheet)
        wb.save('c:\\Users\\Charlie\\Desktop\\packing_list.xlsx')
        packing_button.config(state=tk.DISABLED)
        sheet_name_entry.delete(0, tk.END)

    except FileNotFoundError:
        file_not_found()

    except InvalidFileException:
        no_filename()

    except ValueError:
        no_sheet_name()

        print(counter_panmure)


def make_delivery_list():
    # get method for filename entry
    get_file = filename_entry.get()
    try:
        # loading workbook on local computer c drive using filename
        wb = xl.load_workbook(f'c:\\Users\\Charlie\\Desktop\\{get_file.strip()}.xlsx')

        # working with sheet1 on wb 'workbook'
        sheet = wb['Form responses 3']

        # get method for new sheet name
        sheet_name = sheet_name_entry.get()
        new_sheet = wb.create_sheet("Sheet A", 0)
        new_sheet.title = sheet_name
        sheet_name = wb.active

        # deleting columns so that columns required are left for new file
        sheet.delete_cols(1, 7)
        sheet.delete_cols(9, 4)

        # updating Column Names
        sheet['A1'].value = "First Name"
        sheet['C1'].value = "Street Number"
        sheet['G1'].value = "Delivery Instructions"
        sheet['H1'].value = "Total"

        # calculate total number of rows and columns in source excel file
        max_rows = sheet.max_row
        max_columns = sheet.max_column

        # setting variables for loop
        bold_font = Font(name='Arial', size=14, bold=True)
        # setting font for text in whole sheet except top row
        cell_font = Font(name='Calibri', size=16)
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
        col_clendon = PatternFill(fgColor='edc0b4', fill_type='solid')
        col_henderson = PatternFill(fgColor='eddcb4', fill_type='solid')
        col_howick = PatternFill(fgColor='dbedb4', fill_type='solid')
        col_karaka = PatternFill(fgColor='84b07f', fill_type='solid')
        col_manukau = PatternFill(fgColor='7fb099', fill_type='solid')
        col_manurewa = PatternFill(fgColor='98edc5', fill_type='solid')
        col_meadowbank = PatternFill(fgColor='77d1c8', fill_type='solid')
        col_onehunga = PatternFill(fgColor='87d0ed', fill_type='solid')
        col_otahuhu = PatternFill(fgColor='bebdf2', fill_type='solid')
        col_waiotaiki = PatternFill(fgColor='9292a6', fill_type='solid')
        col_wattle = PatternFill(fgColor='cebad6', fill_type='solid')
        # setting color for 'Total' column
        # totals_color = PatternFill(fgColor='fac2b4', fill_type='solid')

        # copying the cell values from source excel file to destination excel file
        for i in range(1, max_rows + 1):
            for j in range(1, max_columns + 1):
                # reading cell value from source excel file
                c = sheet.cell(row=i, column=j)

                # writing the read value to destination excel file
                new_sheet.cell(row=i, column=j).value = c.value
                new_sheet.cell(row=i, column=j).font = cell_font

                for row in sheet_name.iter_rows(max_row=max_rows, max_col=max_columns):
                    for cell in row:
                        if cell.value == 'Panmure':
                            cell.fill = col_panmure
                            new_sheet.cell(row=i, column=8).fill = col_panmure
                        if cell.value == 'Clendon Park':
                            cell.fill = col_clendon
                            new_sheet.cell(row=i, column=8).fill = col_clendon
                        if cell.value == 'Point England':
                            cell.fill = col_ptengland
                            new_sheet.cell(row=i, column=8).fill = col_ptengland
                        if cell.value == 'Glen Innes':
                            cell.fill = col_gi
                            new_sheet.cell(row=i, column=8).fill = col_gi
                        if cell.value == 'St Johns':
                            cell.fill = col_stjohns
                            new_sheet.cell(row=i, column=8).fill = col_stjohns
                        if cell.value == 'Glendowie':
                            cell.fill = col_glendowie
                            new_sheet.cell(row=i, column=8).fill = col_glendowie
                        if cell.value == 'Mt Wellington':
                            cell.fill = col_mtwell
                            new_sheet.cell(row=i, column=8).fill = col_mtwell
                        if cell.value == 'Greenlane':
                            cell.fill = col_greenlane
                            new_sheet.cell(row=i, column=8).fill = col_greenlane
                        if cell.value == 'Mangere':
                            cell.fill = col_mangere
                            new_sheet.cell(row=i, column=8).fill = col_mangere
                        if cell.value == 'Pakuranga':
                            cell.fill = col_pakuranga
                            new_sheet.cell(row=i, column=8).fill = col_pakuranga
                        if cell.value == 'Henderson':
                            cell.fill = col_henderson
                            new_sheet.cell(row=i, column=8).fill = col_henderson
                        if cell.value == 'Howick':
                            cell.fill = col_howick
                            new_sheet.cell(row=i, column=8).fill = col_howick
                        if cell.value == 'Karaka':
                            cell.fill = col_karaka
                            new_sheet.cell(row=i, column=8).fill = col_karaka
                        if cell.value == 'Manukau':
                            cell.fill = col_manukau
                            new_sheet.cell(row=i, column=8).fill = col_manukau
                        if cell.value == 'Manurewa':
                            cell.fill = col_manurewa
                            new_sheet.cell(row=i, column=8).fill = col_manurewa
                        if cell.value == 'Meadowbank':
                            cell.fill = col_meadowbank
                            new_sheet.cell(row=i, column=8).fill = col_meadowbank
                        if cell.value == 'Onehunga':
                            cell.fill = col_meadowbank
                            new_sheet.cell(row=i, column=8).fill = col_onehunga
                        if cell.value == 'Otahuhu':
                            cell.fill = col_otahuhu
                            new_sheet.cell(row=i, column=8).fill = col_otahuhu
                        if cell.value == 'Waiotaiki Bay':
                            cell.fill = col_waiotaiki
                            new_sheet.cell(row=i, column=8).fill = col_waiotaiki
                        if cell.value == 'Wattle Downs':
                            cell.fill = col_wattle
                            new_sheet.cell(row=i, column=8).fill = col_wattle

                # making row 1 bold font
                new_sheet.cell(row=1, column=i).font = bold_font
                # text alignment for all rows
                new_sheet.cell(row=i, column=j).alignment = horizon_center
                # setting all row height to 30
                sheet_name.row_dimensions[i].height = 40
                # setting borders for all cells
                new_sheet.cell(row=i, column=j).border = border
                # wrapping text on columns 7 & 9
                new_sheet.cell(row=i, column=7).alignment = wrap_text
                new_sheet.cell(row=i, column=9).alignment = wrap_text
                # setting 'Totals' color column to red and bold font
                # new_sheet.cell(row=i, column=8).fill = totals_color
                new_sheet.cell(row=i, column=8).font = bold_font
                # setting specific column widths
                sheet_name.column_dimensions['A'].width = 18
                sheet_name.column_dimensions['B'].width = 18
                sheet_name.column_dimensions['C'].width = 20
                sheet_name.column_dimensions['D'].width = 30
                sheet_name.column_dimensions['E'].width = 25
                sheet_name.column_dimensions['F'].width = 25
                sheet_name.column_dimensions['G'].width = 45
                sheet_name.column_dimensions['H'].width = 12
                sheet_name.column_dimensions['I'].width = 45

        # saving new worksheet to desktop with name packing_list
        wb.remove(sheet)
        wb.save('c:\\Users\\Charlie\\Desktop\\delivery_list.xlsx')
        delivery_button.config(state=tk.DISABLED)
        sheet_name_entry.delete(0, tk.END)

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