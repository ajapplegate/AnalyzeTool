from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlrd
import os

def sort_by_value(d):
    items = d.items()
    backitems = [[str(v[1]), str(v[0])] for v in items]
    backitems.sort(reverse=True)
    return backitems

my_filetypes = [("excel files", "*.xls"), ("excel files", "*.xlsx"), ("all files", "*.*")]

def import_excel():
    import_excel_name = filedialog.askopenfilename(initialdir=os.getcwd(), 
                                                   title="Please select file:",
                                                   filetypes=my_filetypes)
                                                   
    data = xlrd.open_workbook(import_excel_name)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    if ncols == 0 | nrows == 0:
        messagebox.showerror("File format error", "File is empty！")
        return
    if ncols % 2 == 1:
        messagebox.showerror("Wrong file format", "Excel must be an even column！")
        return
    mdict = {}
    for i in range(nrows):
        for j in range(ncols):
            cell = table.cell(i, j).value
            if cell != '':
                mdict[cell] = mdict.get(cell, 0) + 1
    global sortedlist
    sortedlist = sort_by_value(mdict)
    listbox.delete(0, END)
    for item in sortedlist:
        listbox.insert(END, item[1]+' : '+item[0])
    listbox.pack(side=LEFT, fill=BOTH)
    scrollbar.config(command=listbox.yview)


def save_result():
    save_text_name = filedialog.asksaveasfilename(initialdir="/", title="Save file",
                                                  filetypes=(("text files", "*.txt"), ("all files", "*.*")))
    if save_text_name == '':
        return
    if not save_text_name.endswith('.txt'):
        save_text_name += '.txt'
    f = open(save_text_name, W)
    content = ''
    for item in sortedlist:
        content += item[1]+','+item[0] + '\r\n'
    f.write(content)


def show_help():
    messagebox.showinfo("Instructions for use", "Import statistics excel document，Count the number of occurrences of each cell for sheet 0，Reverse order，The result is displayed in the window，It can also be saved to a specified file.")

def show_about():
    messagebox.showinfo("on", "This code was adapted by Anna. If you need help with it, well so do I.")

def init_menu():
    menu_bar = Menu(root)
    file_menu = Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Import excel", command=import_excel)
    file_menu.add_command(label="Save result to:", command=save_result)
    file_menu.add_separator()
    file_menu.add_command(label="drop out", command=root.quit)
    menu_bar.add_cascade(label="operating", menu=file_menu)

    help_menu = Menu(menu_bar, tearoff=0)
    help_menu.add_command(label="Instructions for use", command=show_help)
    help_menu.add_separator()
    help_menu.add_command(label="on", command=show_about)
    menu_bar.add_cascade(label="Help", menu=help_menu)
    root.config(menu=menu_bar)


root = Tk()
init_menu()
scrollbar = Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)
listbox = Listbox(root, yscrollcommand=scrollbar.set)
root.geometry('800x600+500+200')
root.title('Excel analysis tool')
mainloop()
