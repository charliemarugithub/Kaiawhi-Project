import openpyxl as xl
import tkinter as tk
from tkinter import ttk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.borders import BORDER_THICK
from functions import file_not_found, no_filename, no_sheet_name
from functions import no_destination_file, packing_report_generated, delivery_report_generated
from openpyxl.utils.exceptions import InvalidFileException
from collections import defaultdict, Counter
import os

# creating instance of TK class
root = tk.Tk()
# creating main form window and packing it
main_form = tk.Canvas(root, width=500, height=500)
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
main_form.create_window(250, 190, window=sheet_name_entry, width=350, height=25)
'''
# creating entry for destination file
destination_label = tk.Label(root, text='Enter Destination file path name: ', bg='#7289f2', font="Helvetica 16")
main_form.create_window(250, 230, window=destination_label)

# creating entry for filename destination $ placing it in root
destination_entry = tk.Entry(root, font="Helvetica, 16")
main_form.create_window(250, 280, window=destination_entry, width=350, height=25)
'''

def make_packing_list():
    # get method for filename entry
    get_file = filename_entry.get()
    # get_destination = destination_entry.get()
    '''
    # check if destination file exists
    if not os.path.isfile(f'{get_destination.strip()}.xlsx'):
        no_destination_file()
    '''

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
        sheet.insert_cols(1)
        for cell in sheet['D:D']:
            sheet.cell(row=cell.row, column=1, value=cell.value)
        sheet.delete_cols(4)
        sheet.delete_cols(11)
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

        # copying the cell values from source excel file to destination excel file
        for i in range(1, max_rows):
            for j in range(1, max_columns):
                # reading cell value from source excel file
                c = sheet.cell(row=i, column=j)

                # writing the read value to destination excel file
                sheet_name.cell(row=i, column=j).value = c.value
                sheet_name.cell(row=i, column=j).font = cell_font
                # creating dictionary of Suburb and Totals

                for row in sheet_name.iter_rows(max_row=max_rows, max_col=max_columns):
                    for cell in row:
                        if cell.value == 'Panmure':
                            sheet_name.cell(row=i, column=j).fill = col_panmure
                        if cell.value == 'Clendon Park':
                            sheet_name.cell(row=i, column=j).fill = col_clendon
                        if cell.value == 'Point England':
                            sheet_name.cell(row=i, column=j).fill = col_ptengland
                        if cell.value == 'Glen Innes':
                            sheet_name.cell(row=i, column=j).fill = col_gi
                        if cell.value == 'St Johns':
                            sheet_name.cell(row=i, column=j).fill = col_stjohns
                        if cell.value == 'Glendowie':
                            sheet_name.cell(row=i, column=j).fill = col_glendowie
                        if cell.value == 'Mt Wellington':
                            sheet_name.cell(row=i, column=j).fill = col_mtwell
                        if cell.value == 'Greenlane':
                            sheet_name.cell(row=i, column=j).fill = col_greenlane
                        if cell.value == 'Mangere':
                            sheet_name.cell(row=i, column=j).fill = col_mangere
                        if cell.value == 'Pakuranga':
                            sheet_name.cell(row=i, column=j).fill = col_pakuranga
                        if cell.value == 'Henderson':
                            sheet_name.cell(row=i, column=j).fill = col_henderson
                        if cell.value == 'Howick':
                            sheet_name.cell(row=i, column=j).fill = col_howick
                        if cell.value == 'Karaka':
                            sheet_name.cell(row=i, column=j).fill = col_karaka
                        if cell.value == 'Manukau':
                            sheet_name.cell(row=i, column=j).fill = col_manukau
                        if cell.value == 'Manurewa':
                            sheet_name.cell(row=i, column=j).fill = col_manurewa
                        if cell.value == 'Meadowbank':
                            sheet_name.cell(row=i, column=j).fill = col_meadowbank
                        if cell.value == 'Onehunga':
                            sheet_name.cell(row=i, column=j).fill = col_onehunga
                        if cell.value == 'Otahuhu':
                            sheet_name.cell(row=i, column=j).fill = col_otahuhu
                        if cell.value == 'Waiotaiki Bay':
                            sheet_name.cell(row=i, column=j).fill = col_waiotaiki
                        if cell.value == 'Wattle Downs':
                            sheet_name.cell(row=i, column=j).fill = col_wattle

                # making row 1 bold font
                sheet_name.cell(row=1, column=i).font = bold_font
                # text alignment for all rows
                sheet_name.cell(row=i, column=j).alignment = horizon_center
                # setting all row height to 30
                sheet_name.row_dimensions[i].height = 45
                # setting borders for all cells
                sheet_name.cell(row=i, column=j).border = border
                # wrapping text on columns 8-10
                sheet_name.cell(row=i, column=8).alignment = wrap_text
                sheet_name.cell(row=i, column=9).alignment = wrap_text
                sheet_name.cell(row=i, column=10).alignment = wrap_text
                # Totals column made bold
                sheet_name.cell(row=i, column=5).font = bold_font
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

        wb.remove(sheet)
        # create 2nd sheet to copy over suburb and totals
        box_name = 'Boxes'
        box_sheet = wb.create_sheet("Sheet B", 1)
        box_sheet.title = box_name
        box_name = wb.active

        # creating 2 lists to take suburbs and totals
        suburbs_list = []
        totals_list = []

        # creating dictionary to collect 2 lists above
        sub_and_totals = defaultdict(list)

        # iterating over suburbs and appending to suburbs list
        for cell in sheet['A:A']:
            # writing values to new sheet
            box_sheet.cell(row=cell.row, column=1, value=cell.value)
            suburbs_list.append(cell.value)
        # iterating over totals and appending to totals list
        for cell in sheet['E:E']:
            # writing values to new sheet
            box_sheet.cell(row=cell.row, column=2, value=cell.value)
            totals_list.append(cell.value)

        # removing row 1 as not needed for dictionary
        box_sheet.delete_rows(1)

        # iterating over both lists to append to dictionary without duplicate
        # keys (zip)  and appending values that belong to the same key
        for i, j in zip(suburbs_list, totals_list):
            sub_and_totals[i].append(j)

        del sub_and_totals['Suburb']

        print(sub_and_totals)
        # creating variable dict_tables to take dict values
        dict_values = sub_and_totals.values()
        # counting how often duplicates values are in this dict
        for counter in dict_values:
            frequency = Counter(counter)
            print(frequency)


        '''
        # creating another dictionary to count keys frequency
        # this counts keys only, not values
        # Do I need this right now? No, comment out for now

        frequency = {}
        for item in sub_and_totals:
            if item in frequency:
                frequency[item] += 1
            else:
                frequency[item] = 1

        print(frequency)
        '''

        # saving new worksheet to desktop with name packing_list
        wb.save('c:\\Users\\Charlie\\Desktop\\packing_list.xlsx')
        packing_button.config(state=tk.DISABLED)
        sheet_name_entry.delete(0, tk.END)
        # destination_entry.delete(9, tk.END)
        packing_report_generated()

    except FileNotFoundError:
        file_not_found()

    except InvalidFileException:
        no_filename()

    except ValueError:
        no_sheet_name()


def make_delivery_list():
    get_file = filename_entry.get()
    # get_destination = destination_entry.get()
    '''
    # check if destination field is empty
    if get_destination == '':
        no_destination_file()
    else:
    '''
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
        sheet.delete_cols(1, 7)
        sheet.delete_cols(9, 4)
        sheet.insert_cols(1)
        for cell in sheet['F:F']:
            sheet.cell(row=cell.row, column=1, value=cell.value)
        sheet.delete_cols(6)
        sheet.delete_cols(10)

        # updating Column Names
        sheet['B1'].value = "First Name"
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

        # copying the cell values from source excel file to destination excel file
        for i in range(1, max_rows):
            for j in range(1, max_columns):
                # reading cell value from source excel file
                c = sheet.cell(row=i, column=j)

                # writing the read value to destination excel file
                sheet_name.cell(row=i, column=j).value = c.value
                sheet_name.cell(row=i, column=j).font = cell_font

                for row in sheet_name.iter_rows(max_row=max_rows, max_col=max_columns):
                    for cell in row:
                        if cell.value == 'Panmure':
                            sheet_name.cell(row=i, column=j).fill = col_panmure
                        if cell.value == 'Clendon Park':
                            sheet_name.cell(row=i, column=j).fill = col_clendon
                        if cell.value == 'Point England':
                            sheet_name.cell(row=i, column=j).fill = col_ptengland
                        if cell.value == 'Glen Innes':
                            sheet_name.cell(row=i, column=j).fill = col_gi
                        if cell.value == 'St Johns':
                            sheet_name.cell(row=i, column=j).fill = col_stjohns
                        if cell.value == 'Glendowie':
                            sheet_name.cell(row=i, column=j).fill = col_glendowie
                        if cell.value == 'Mt Wellington':
                            sheet_name.cell(row=i, column=j).fill = col_mtwell
                        if cell.value == 'Greenlane':
                            sheet_name.cell(row=i, column=j).fill = col_greenlane
                        if cell.value == 'Mangere':
                            sheet_name.cell(row=i, column=j).fill = col_mangere
                        if cell.value == 'Pakuranga':
                            sheet_name.cell(row=i, column=j).fill = col_pakuranga
                        if cell.value == 'Henderson':
                            sheet_name.cell(row=i, column=j).fill = col_henderson
                        if cell.value == 'Howick':
                            sheet_name.cell(row=i, column=j).fill = col_howick
                        if cell.value == 'Karaka':
                            sheet_name.cell(row=i, column=j).fill = col_karaka
                        if cell.value == 'Manukau':
                            sheet_name.cell(row=i, column=j).fill = col_manukau
                        if cell.value == 'Manurewa':
                            sheet_name.cell(row=i, column=j).fill = col_manurewa
                        if cell.value == 'Meadowbank':
                            sheet_name.cell(row=i, column=j).fill = col_meadowbank
                        if cell.value == 'Onehunga':
                            sheet_name.cell(row=i, column=j).fill = col_onehunga
                        if cell.value == 'Otahuhu':
                            sheet_name.cell(row=i, column=j).fill = col_otahuhu
                        if cell.value == 'Waiotaiki Bay':
                            sheet_name.cell(row=i, column=j).fill = col_waiotaiki
                        if cell.value == 'Wattle Downs':
                            sheet_name.cell(row=i, column=j).fill = col_wattle

                # making row 1 bold font
                sheet_name.cell(row=1, column=i).font = bold_font
                # text alignment for all rows
                sheet_name.cell(row=i, column=j).alignment = horizon_center
                # setting all row height to 30
                sheet_name.row_dimensions[i].height = 65
                # setting borders for all cells
                sheet_name.cell(row=i, column=j).border = border
                # wrapping text on columns 7 & 9
                sheet_name.cell(row=i, column=7).alignment = wrap_text
                sheet_name.cell(row=i, column=9).alignment = wrap_text
                # Totals column made bold
                sheet_name.cell(row=i, column=8).font = bold_font
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
        # destination_entry.delete(9, tk.END)
        delivery_report_generated()

    except FileNotFoundError:
        file_not_found()

    except InvalidFileException:
        no_filename()

    except ValueError:
        no_sheet_name()


def clear_all():
    delivery_button.config(state=tk.ACTIVE)
    packing_button.config(state=tk.ACTIVE)
    sheet_name_entry.delete(0, tk.END)
    filename_entry.delete(0, tk.END)
    # destination_entry.delete(0, tk.END)


packing_button = ttk.Button(text='Packing List', command=make_packing_list)
main_form.create_window(150, 360, window=packing_button, height=50, width=150)

delivery_button = ttk.Button(text='Delivery List', command=make_delivery_list)
main_form.create_window(350, 360, window=delivery_button, height=50, width=150)

clear_button = ttk.Button(text='CLEAR ALL', command=clear_all)
main_form.create_window(250, 450, window=clear_button, height=50, width=150)

root.mainloop()
