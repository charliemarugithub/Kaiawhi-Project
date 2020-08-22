from tkinter import messagebox


def file_not_found():
    messagebox.showinfo("File Name Error!", "File Not Found, Type Correct Source File Name.")


def no_filename():
    messagebox.showinfo("File Name Cannot Be Empty!", "You Must Enter A File Name.")


def no_sheet_name():
    messagebox.showinfo("No Sheet Name Listed!", "You Must Enter A Sheet Name.")


def packing_report_generated():
    messagebox.showinfo("Report!", "Packing List Now Completed.")


def delivery_report_generated():
    messagebox.showinfo("Report!", "Delivery List Now Completed.")


def no_destination_file():
    messagebox.showinfo("No Destination!", "Enter File Destination Path.")