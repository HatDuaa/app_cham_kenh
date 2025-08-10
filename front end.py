from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter.filedialog import Open
from tkinter import scrolledtext
from threading import Thread
import tkinter
import back_end as back_end


def thongKe_window():
    if var_ThongKe_SDD.get():
        txtThongKe = back_end.thongKe_SDD()
    elif var_ThongKe_CNCC.get():
        txtThongKe = back_end.thongKe_CNCC()

    windowScrolledText = Tk()
    windowScrolledText.geometry('700x500')
    txt = scrolledtext.ScrolledText(windowScrolledText, width=200, height=50)
    txt.insert(INSERT, txtThongKe)
    txt.pack()
    messagebox.showinfo('', 'Hoàn thành thống kê')
    windowScrolledText.mainloop()

def clickNew():
    global filename1
    ftypes = [('Excel file', '*.xlsx'), ('All files', '*')]
    dlg = Open(filetypes=ftypes)
    filename1 = dlg.show()
    lblFile.configure(text=filename1)
    print(filename1)

def thread_mess(text):
    thread = Thread(target=messagebox.showinfo('', text))
    thread.start()

def handleButton1():
    global load1
    global sheet1
    global sheet2
    global load3
    global sheet3
    global load4
    global sheet4


    if filename1 == '':
        messagebox.showinfo('', 'Chưa chọn file!!')
    else:
        load1 = openpyxl.load_workbook(filename1)
        sheet1 = load1['Sheet1']
        if var_RandomCNCC.get():
            load2 = openpyxl.load_workbook('Bang cham kenh danh cho tre SDD .xlsx')
            sheet2 = load2['Sheet1']
            thread_mess('Đang điền CN_CC')
            randomCNCC()

        if var_SDD.get():
            load2 = openpyxl.load_workbook('Bang cham kenh danh cho tre SDD .xlsx')
            sheet2 = load2['Sheet1']
            thread_mess('Đang chấm SDD')
            chamSDD()

        if var_SDDII.get():
            if var_SDD.get() or var_CNCC.get() or var_ThongKe_SDD.get():
                messagebox.showinfo('', 'Không thể thực hiện chấm SDD độ II cùng các thao tác khác')
            else:
                load3 = openpyxl.load_workbook('Bang cham kenh SDD do II.xlsx')
                sheet3 = load3['Sheet1']
                thread_mess('Đang chấm SDD độ II')
                chamSDDII()

        if var_CNCC.get():
            load4 = openpyxl.load_workbook('Bang tra CNCC.xlsx')
            sheet4 = load4['Sheet1']
            thread_mess('Đang chấm CN/CC')
            chamCNCC()

        if var_ThongKe_SDD.get():
            thread_mess('Đang thống kê SDD')
            thongKe_window()

        if var_ThongKe_CNCC.get():
            thread_mess('Đang thống kê CN/CC')
            thongKe_window()


    #messagebox.showinfo('hello world', 'Hoàn thành')
    return

def handleButton2():
    thread =Thread(target=handleButton1)
    thread.start()



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
