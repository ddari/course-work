import subprocess
import platform
import argparse

def ping_ip(ip_address):
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
        reply = subprocess.run(['ping','-n 1', ip_address],
                               stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE,
                               encoding='utf-8')
    else: 
        reply = subprocess.run(['ping', '-c', '1', '-n', ip_address],
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE,
                           encoding='utf-8')
    if reply.returncode == 0:
        return True, reply.stdout
    else:
        return False, reply.stderr

print(ping_ip('8.8.8.8'))
print(ping_ip('a'))

parser = argparse.ArgumentParser(description='Ping script')

parser.add_argument('-a', action="store", dest="ip")
parser.add_argument('-c', action="store", dest="count", default=2, type=int)

args = parser.parse_args()
print(args)

rc, message = ping_ip(args.ip)
print(message)
