import subprocess
import platform
import argparse
import re
import openpyxl

# import pythonping as pig

def ping_ip(ip_address: str, count=1) -> (bool, str):
    """
    Ping IP address and return tuple:
    On success:
        * True
        * command output (stdout)
    On failure:
        * False
        * error output (stderr)
    """
    # для разных систем разные кодировки
    if platform.system().lower() == "windows" :
        reply = subprocess.run(['ping','-n ',str(count), ip_address],
                               stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE,
                               encoding='cp866')
    else: 
        reply = subprocess.run(['ping', '-c', str(count), '-n', ip_address],
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE,
                           encoding='utf-8')
    # print(reply.stdout)
    pink_ms = re.search(r'=\d*[ms,мс]', reply.stdout)
     # reply.returncode
    if pink_ms is not None:
        return True, reply.stdout[pink_ms.start() + 1:pink_ms.end() - 1]
    else:
        return False, ''
    
def list_from_exel(file, paper) -> list:
    if paper is None:
        paper = 'Лист1'
    l = []
    k = 1
    wb = openpyxl.load_workbook(filename=file, read_only=True)
    sheet = wb[paper]
    a = sheet.cell(row=k, column=1).value
    while a is not None:
        l.append(a)
        k += 1
        a = sheet.cell(row=k, column=1).value
    return l

def list_to_ping_list(l: list) -> [(str, str)]:
    ll = []
    for i in l:
        t = ping_ip(i)
        if t[0]:
            ll.append((i, t[1]))
        else:
            ll.append((i, 'Ошибка соединения'))
    return ll

def print_single_ip(args):
    go, ms = ping_ip(args.ip, args.count)
    if go:
        print(f'Ping {ms} мс')
    else:
        print('Ответ не вернулся')


def print_many_ip(args):
    ip_list = list_from_exel(args.file, args.list)
    ping_list = list_to_ping_list(ip_list)
    for i in range(len(ping_list)):
        print(f'{i + 1}. {ping_list[i][0]} {ping_list[i][1]}')


def give_file(args):
    ip_list = list_from_exel(args.file, args.list)
    ping_list = list_to_ping_list(ip_list)
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Лист1'
    sheet.cell(row=1, column=1).value = 'IP'
    sheet.cell(row=1, column=2).value = 'Ping'
    for i in range(len(ping_list)):
        sheet.cell(row=i + 2, column=1).value = ping_list[i][0]
        sheet.cell(row=i + 2, column=2).value = ping_list[i][1]
    print('Документ успешно сформирован')
    wb.save(args.out + '.xlsx')


def repeat(args):
    ip_list = list_from_exel(args.repeat + '.xlsx', None)
    ping_list = list_to_ping_list(ip_list)
    wb = openpyxl.load_workbook(filename=args.repeat + '.xlsx')
    sheet = wb['Лист1']
    k = 2
    a = sheet.cell(row=k, column=2).value
    while a is not None:
        column = 2
        b = sheet.cell(row=k, column=column).value
        while b is not None:
            column += 1
            b = sheet.cell(row=k, column=column).value
        sheet.cell(row=k, column=column).value = ping_list[k - 2][1]
        k += 1
        a = sheet.cell(row=k, column=1).value
    wb.save(args.repeat + '.xlsx')


def find_dif(args):
    def sred_arf(lt: list):
        sumar = 0
        try:
            for i in lt:
                sumar += int(i)
        except:
            return 'Ответ не вернулся'
        return float(sumar) / len(lt)

    ip_list = list_from_exel(args.dif + '.xlsx', None)
    ping_list = list_to_ping_list(ip_list)
    wb = openpyxl.load_workbook(filename=args.dif + '.xlsx')
    sheet = wb['Лист1']
    k = 2
    column = 3
    a = sheet.cell(row=k, column=2).value
    while a is not None:
        column = 2
        l = []
        b = sheet.cell(row=k, column=column).value
        while b is not None:
            try:
                l.append(int(b))
            except:
                l.append(b)
            column += 1
            b = sheet.cell(row=k, column=column).value
        l.append(ping_list[k - 2][1])
        sheet.cell(row=k, column=column + 1).value = sred_arf(l)
        sheet.cell(row=k, column=column).value = ping_list[k - 2][1]
        k += 1
        a = sheet.cell(row=k, column=1).value
    sheet.cell(row=1, column=column + 1).value = "Среднее арифметическое"
    wb.save(args.dif + '.xlsx')

if __name__ == '__main__':
    # для запуска из командной строки с аргументами
    parser = argparse.ArgumentParser(description='Ping script')
    parser.add_argument('-a', action="store", dest="ip", type=str)  # '140.82.121.4'
    parser.add_argument('-c', action="store", dest="count", default=1, type=int)
    parser.add_argument('-f', action="store", dest="file", type=str)
    parser.add_argument('-l', action="store", dest="list", type=str)
    parser.add_argument('-o', action="store", dest="out", type=str)
    parser.add_argument('-r', action="store", dest="repeat", type=str)
    parser.add_argument('-d', action="store", dest="dif", type=str)
    
    args = parser.parse_args()
    if args.out:
        give_file(args)
    if args.ip:
        print_single_ip(args)
    if args.file:
        if '.csv' in args.file or '.xlsx' in args.file:
            print_single_ip(args)
       else:
            print('Недопустимый формат файла')
   if args.repeat:
        repeat(args)
        try:
            repeat(args)
        except:
            print('файл не найден или его формат неверный')

    if args.dif:
        find_dif(args)
        try:
            find_dif(args)
        except:
            print('файл не найден или его формат неверный')
