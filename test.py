import time

import openpyxl
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter.filedialog import Open
from tkinter import scrolledtext
from threading import Thread
import tkinter

def clickNew():
    global filename1
    ftypes = [('Excel file', '*.xlsx'), ('All files', '*')]
    dlg = Open(filetypes=ftypes)
    filename1 = dlg.show()
    lblFile.configure(text=filename1)
    print(filename1)

def thread_mess(text):
    messagebox._show('', 'text')

def handleButton1():
    messagebox._show('', 'text')
    time.sleep(3)



def handleButton2():
    thread =Thread(target=handleButton1)
    thread.start()
    print('a')


window = Tk()
window.title('Nguyễn Lộc')
window.geometry("700x500")

# Thêm menu bar, Lấy đường dẫn
filename1 = ''
menu = Menu(window)
new_item = Menu(menu, tearoff=0)
new_item.add_command(label='New', command=clickNew)
menu.add_cascade(label='File', menu=new_item)
window.config(menu=menu)

# Thêm label
hello = "Xin Chào!!"
lblHello = tkinter.Label(window, text=hello, fg="red", font=("Times new roman", 20))
lblHello.place(x=0, y=0)

lblFile = Label(window, text='', font=("Times new roman", 13))
lblFile.place(x=0, y=40)

# Thêm combobox
'''combo = Combobox(window)
combo['values'] = ("Chấm SDD", "Chấm SDD độ I-II", "Chấm CN/CC", "Thống kê")
combo.current(0)
combo.grid(column=0, row=2)'''

# Thêm button
btnChon = Button(window, text="Chọn", command=handleButton2)
btnChon.place(x=200, y=200)

# Thêm checkbox
var_SDD = BooleanVar()
var_SDDII = BooleanVar()
var_CNCC = BooleanVar()
var_ThongKe_SDD = BooleanVar()
var_ThongKe_CNCC = BooleanVar()
var_RandomCNCC = BooleanVar()

cb_SDD = Checkbutton(window, text="Chấm SDD", variable=var_SDD)
cb_SDDII = Checkbutton(window, text="Chấm SDD độ II", variable=var_SDDII)
cb_ChamCNCC = Checkbutton(window, text="Chấm CN/CC", variable=var_CNCC)
cb_ThongKe_SDD = Checkbutton(window, text="Thống kê SDD", variable=var_ThongKe_SDD)
cb_ThongKe_CNCC = Checkbutton(window, text="Thống kê CN/CC", variable=var_ThongKe_CNCC)
cb_RandomCNCC = Checkbutton(window, text="Điền CN_CC", variable=var_RandomCNCC)

cb_SDD.place(x=50, y=100)
cb_SDDII.place(x=300, y=100)
cb_RandomCNCC.place(x=550, y=100)
cb_ChamCNCC.place(x=50, y=150)
cb_ThongKe_SDD.place(x=300, y=150)
cb_ThongKe_CNCC.place(x=550, y=150)



window.mainloop()
