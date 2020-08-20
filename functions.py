from tkinter import messagebox


def file_not_found():
    messagebox.showinfo("File Name Error!", "File Not Found, Type Correct File Name.")


def no_filename():
    messagebox.showinfo("File Name Cannot Be Empty!", "You Must Enter A File Name.")


def no_sheet_name():
    messagebox.showinfo("No Sheet Name Listed!", "You Must Enter A Sheet Name.")


def report_generating():
    messagebox.showinfo("Report!", "Your Report Is Being Generated.")


def permission_error():
    messagebox.showinfo("This file is open somewhere, please close it and start again!")
