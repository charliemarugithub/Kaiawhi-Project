from tkinter import messagebox, DISABLED


def file_not_found():
    messagebox.showinfo("File Name Error!", "File Not Found, Type Correct File Name.")


def no_filename():
    messagebox.showinfo("File Name Cannot Be Empty!", "You Must Enter A File Name.")


def no_sheet_name():
    messagebox.showinfo("No Sheet Name Listed!", "You Must Enter Sheet Name.")


def report_generating():
    messagebox.showinfo("Report!", "Your Report Is Being Generated.")


# def disable_packing_btn():
# packing_button.config(state=DISABLED)


# def disable_delivery_btn():
# delivery_button.config(state=DISABLED)
