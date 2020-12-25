from tkinter import *
from tkinter import ttk, filedialog, messagebox
from ping_checker import *
import numpy as n
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt


def clean_res():  # очистка полей значений
    XList.delete(0, END)
    YList.delete(0, END)


def save_file():  # сохранение файла
    if '.xlsx' not in OutStr.get():
        lace = filedialog.asksaveasfilename(defaultextension='.xlsx')
        OutStr.set(lace)
    else:
        lace = OutStr.get()

    l = XList.get(0, END)
    ping_list = list_to_ping_list(l)
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Лист1'
    sheet.cell(row=1, column=1).value = 'IP'
    sheet.cell(row=1, column=2).value = 'Ping'
    for i in range(len(ping_list)):
        sheet.cell(row=i + 2, column=1).value = ping_list[i][0]
        sheet.cell(row=i + 2, column=2).value = ping_list[i][1]
    print('Документ успешно сформирован')
    wb.save(lace)
    print(l)


def load_file():  # загрузка из файла
    if '.xlsx' not in InStr.get():
        lace = filedialog.askopenfilename(defaultextension='.xlsx')
        InStr.set(lace)
    else:
        lace = InStr.get()
    for i in list_to_ping_list(list_from_exel(lace[:-5], None)):
        XList.insert(END, i[0])
        YList.insert(END, i[1])


def ping():
    go, ms = ping_ip(IPStr.get(), 1)
    XList.insert(END, IPStr.get())
    if go:
        YList.insert(END, f'Ping {ms} мс')
    else:
        YList.insert(END, 'Ответ не вернулся')


def add_column():
    if '.xlsx' not in InStr.get():
        lace = filedialog.askopenfilename()
        InStr.set(lace)
    else:
        lace = InStr.get()
    s = lace[:-5]
    repeat(s)


def srav():
    if '.xlsx' not in InStr.get():
        lace = filedialog.askopenfilename()
        InStr.set(lace)
    else:
        lace = InStr.get()
    s = lace[:-5]
    find_dif(s)


def build(window):  # постройка окна
    global XList, YList

    style = ttk.Style()
    # available_themes = style.theme_names()
    style.theme_use('vista')  # задаём стиль по фану

    # создание элементов
    lbl = Label(text="IP = ").place(x=30, y=20)
    txt_A = Entry(width=20, textvariable=IPStr).place(x=60, y=20)
    Ping = Button(text="Ping", command=ping).place(x=30, y=50)

    lbl1 = Label(text="File IN = ").place(x=30, y=80)
    txt_I = Entry(width=20, textvariable=InStr).place(x=100, y=80)
    ButLoad1 = Button(text="Выбрать файл", command=load_file).place(x=250, y=80)

    lbl2 = Label(text="File out = ").place(x=30, y=110)
    txt_O = Entry(width=20, textvariable=OutStr).place(x=100, y=110)
    ButSave = Button(text="Сохранить в файл", command=save_file).place(x=250, y=110)

    ButLoad = Button(text="Загрузить из файла", command=load_file).place(x=100, y=50)
    Dif = Button(text="Провести сравнение", command=srav).place(x=300, y=50)
    MakeCol = Button(text="Добавить ещё столбец", command=add_column).place(x=500, y=50)

    # ButQ = Button(ext='Выход', command=quit())

    lb31 = Label(text='IP ')
    lb32 = Label(text='Ответ')

    scrollbar1 = Scrollbar()
    scrollbar2 = Scrollbar()

    XList = Listbox(yscrollcommand=scrollbar1.set, bd=1)
    YList = Listbox(yscrollcommand=scrollbar2.set, bd=1)
    scrollbar1.config(command=XList.yview)  # привязка к листу
    scrollbar2.config(command=YList.yview)

    ButClean = Button(text='ОЧИСТИТЬ', command=clean_res)

    # разместить элементы
    lb31.place(x=40, y=130)
    lb32.place(x=180, y=130)

    XList.place(x=40, y=150, width=80, height=200)
    scrollbar1.place(x=120, y=150, height=200)
    YList.place(x=180, y=150, width=80, height=200)
    scrollbar2.place(x=260, y=150, height=200)

    ButClean.place(x=40, y=400, width=200, height=60)


if __name__ == '__main__':
    window = Tk()
    window.title("Добро пожаловать в Курсач")
    window.geometry('730x500')
    XList, YList = '', ''  # костыль, чтобы не делать через класс
    IPStr, OutStr, InStr = (StringVar() for x in range(3))  # анологично
    build(window)
    window.mainloop()
