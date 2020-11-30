import subprocess
import platform
import argparse
import re

def ping_ip(ip_address, count, wait=120):
    """
    Ping IP address and return tuple:
    On success:
        * True
        * command output (stdout)
    On failure:
        * False
        * error output (stderr)
    """
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
    pink_ms = re.search(r'=\d*[ms,мс]', reply.stdout)
    if pink_ms:
        print()
    if reply.returncode == 0:
        return True, reply.stdout[pink_ms.start() + 1:pink_ms.end() - 1]
    else:
        return False, wait

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Ping script')

    parser.add_argument('-a', action="store", dest="ip", default='140.82.121.4', type=str)
    parser.add_argument('-c', action="store", dest="count", default=4, type=int)
    args = parser.parse_args()

    go, ms = ping_ip(args.ip, args.count)
    if go:
        print(f'Ответ вернулся через {ms} мс')
    else:
        print(f'Ответ не вернулся')
