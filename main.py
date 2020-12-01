import subprocess
import platform
import argparse
import re
import openpyxl

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
    if reply.returncode == 0:
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

if __name__ == '__main__':
    # для запуска из командной строки с аргументами
    parser = argparse.ArgumentParser(description='Ping script')
    parser.add_argument('-a', action="store", dest="ip", type=str)  # '140.82.121.4'
    parser.add_argument('-c', action="store", dest="count", default=1, type=int)
    parser.add_argument('-f', action="store", dest="file", type=str)
    parser.add_argument('-l', action="store", dest="list", type=str)
    args = parser.parse_args()

    if args.ip:
        go, ms = ping_ip(args.ip, args.count)
        if go:
            print(f'Ответ вернулся через {ms} мс')
        else:
            print(f'Ответ не вернулся')
    if args.file:
        if '.csv' in args.file or '.xlsx' in args.file:
            ip_list = list_from_exel(args.file, args.list)
            ping_list = list_to_ping_list(ip_list)
            for i in range(len(ping_list)):
                print(f'{i+1}. {ping_list[i][0]} {ping_list[i][1]}')
