from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlrd


def sort_by_value(d):
    items = d.items()
    backitems = [[str(v[1]), str(v[0])] for v in items]
    backitems.sort(reverse=True)
    return backitems


def import_excel():
    import_excel_name = filedialog.askopenfilename(initialdir="/", title="选择EXCEL文件",
                                                   filetypes=(("excel files", "*.xls"),
                                                              ("excel files", "*.xlsx"),
                                                              ("all files", "*.*"))
                                                   )
    data = xlrd.open_workbook(import_excel_name)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    if ncols == 0 | nrows == 0:
        messagebox.showerror("文件格式错误", "excel为空！")
        return
    if ncols % 2 == 1:
        messagebox.showerror("文件格式错误", "excel必须为偶数列！")
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
    save_text_name = filedialog.asksaveasfilename(initialdir="/", title="保存统计结果",
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
    messagebox.showinfo("操作说明", "导入需统计的EXCEL文档，对sheet0统计每个单元格出现的次数，倒序排列，结果在窗口中显示，亦可保存至指定文件。")


def show_about():
    messagebox.showinfo("关于", "本软件由vince开发，如有疑问请email至vincehe2013@163.com。")


def init_menu():
    menu_bar = Menu(root)
    file_menu = Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="导入EXCEL", command=import_excel)
    file_menu.add_command(label="结果保存至..", command=save_result)
    file_menu.add_separator()
    file_menu.add_command(label="退出", command=root.quit)
    menu_bar.add_cascade(label="操作", menu=file_menu)

    help_menu = Menu(menu_bar, tearoff=0)
    help_menu.add_command(label="使用说明", command=show_help)
    help_menu.add_separator()
    help_menu.add_command(label="关于", command=show_about)
    menu_bar.add_cascade(label="帮助", menu=help_menu)
    root.config(menu=menu_bar)


root = Tk()
init_menu()
scrollbar = Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)
listbox = Listbox(root, yscrollcommand=scrollbar.set)
root.geometry('800x600+500+200')
root.title('EXCEL计数工具V1.0')
mainloop()
