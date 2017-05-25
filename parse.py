#-*-coding: utf-8-*-
#!/usr/bin/python

import xlrd
import glob
from Tkinter import *
from ScrolledText import *
import config as cfg
import paramiko
import time
import ipcalc

# xls columns
class XLS(object):
    SERV = 3
    ADDR = 4
    IP = 12
    VIP = 13
    PORT = 19
    FIRST = 4 # first row
    PING = 'ping -m 1 -c 3 -vpn-instance 112 %s\n'
    SLEEP = 8

    def __setattr__(self, *_):
        pass
XLS = XLS()


def getDataFromFile(fileName):
    with xlrd.open_workbook(fileName) as wb:
        worksheet = wb.sheet_by_index(0)
        data = {}
        for r in range(XLS.FIRST, worksheet.nrows):
            data[r] = {}
            for c in range(0, XLS.PORT + 1):
                data[r][c] = worksheet.cell(r, c).value

        return data

for file in glob.glob("*.xlsx"):
    data = getDataFromFile(file)


# address list for dropmenu
OPTIONS = []
WIDTH = 0
for r in data:
    # get rows with IPs only
    if data[r][XLS.IP] != '':
        addr = str(r + 1) + ': ' + data[r][XLS.SERV] + ' ' + data[r][XLS.ADDR]
        OPTIONS.append(addr)
        if WIDTH == 0:
            init = r
        if WIDTH < len(addr):
            WIDTH = len(addr)


# get IPs
def get(opt):
    global address
    result = re.search(r'(\S+):', opt)
    if result:
        found = result.group(1)
    address = data[int(found) - 1][XLS.ADDR]

    result = re.search(r'([\d\.\/]+)([^\d]+)([\d\.\/]+)([^\d]*)([\d\.\/]*)', data[int(found) - 1][XLS.IP])
    if result:
        str1 = result.group(1)
        str2 = result.group(3)
        str3 = result.group(5)
    if str1 != '':
        IP1.set(str1)
    else:
        IP1.set('x')
    if str2 != '':
        IP2.set(str2)
    else:
        IP2.set('x')
    if str3 != '':
        IP3.set(str3)
    else:
        IP3.set('x')

    result = re.search(r'([\d\.\/]*)([^\d]*)([\d\.\/]*)', data[int(found) - 1][XLS.VIP])
    if result:
        strv1 = result.group(1)
        strv2 = result.group(3)
    if strv1 != '':
        IPV1.set(strv1)
    else:
        IPV1.set('x')
    if strv2 != '':
        IPV2.set(strv2)
    else:
        IPV2.set('x')

    strInt = data[int(found) - 1][XLS.PORT]
    if strInt != '':
        Int.set(strInt)
    else:
        Int.set('x')


def ping_pl(text):
    for line in text.splitlines():
         result = re.search(r'([\d\.]+)% packet loss', line)
         if result:
             break
    return result.group(1)
    
def mac_pl(text):
    for line in text.splitlines():
         result = re.search(r'([0-9a-fA-F]{2}(?::[0-9a-f]{2}){5})', line)
         if result:
             break
    return result.group(1)
    
def pinger():
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(cfg.ssh['host'], username = cfg.ssh['login'], password = cfg.ssh['password'], port = cfg.ssh['port'])

        channel = ssh.invoke_shell()
        channel.settimeout(cfg.ssh['timeout'])

        channel.send('open ' + cfg.telnet['host'] + '\n')
        time.sleep(5)
        output = channel.recv(1024)
        print output,
        channel.send(cfg.telnet['login'] + '\n')
        time.sleep(1)
        output = channel.recv(1024)
        print output,
        channel.send(cfg.telnet['password'] + '\n')
        time.sleep(1)
        output = channel.recv(1024)
        print output,
        channel.send(XLS.PING % str(ip1))
        time.sleep(XLS.SLEEP)
        output = channel.recv(65536)
        print output,
        ip1_ping = ping_pl(output)
        ent.insert(END, 'Основной %s%% потерь\n' % ip1_ping)
        channel.send(XLS.PING % str(ip2))
        time.sleep(XLS.SLEEP)
        output = channel.recv(65536)
        print output,
        ip2_ping = ping_pl(output)
        ent.insert(END, 'Резервный %s%% потерь\n' % ip2_ping)

        if IP3.get() != 'x':
            channel.send(XLS.PING % str(ip3))
            time.sleep(XLS.SLEEP)
            output = channel.recv(65536)
            print output,
            ip3_ping = ping_pl(output)
            ent.insert(END, 'Третий %s%% потерь\n' % ip3_ping)

        if IPV1.get() != 'x':
            channel.send(XLS.PING % str(ipv1))
            time.sleep(XLS.SLEEP)
            output = channel.recv(65536)
            print output,
            ipv1_ping = ping_pl(output)
            ent.insert(END, 'VipNet1 %s%% потерь\n' % ipv1_ping)

        if IPV2.get() != 'x':
            channel.send(XLS.PING % str(ipv2))
            time.sleep(XLS.SLEEP)
            output = channel.recv(65536)
            print output,
            ipv2_ping = ping_pl(output)
            ent.insert(END, 'VipNet2 %s%% потерь\n' % ipv2_ping)

        if Int.get() != 'x':
            ip = ''
            if float(ip1_ping) < 100:
                ip = str(ip1)
            elif float(ip2_ping) < 100:
                ip = str(ip2)
            if ip != '':
                channel.send('telnet vpn-instance 112 ' + ip + '\n')
                time.sleep(5)
                output = channel.recv(1024)
                print output,
                channel.send(cfg.edds['login'] + '\n')
                time.sleep(1)
                output = channel.recv(1024)
                print output,
                channel.send(cfg.edds['password'] + '\n')
                time.sleep(3)
                output = channel.recv(1024)
                print output,
##                channel.send('show interfaces descriptions ' + Int.get() + '\n')
##                time.sleep(3)
##                output = channel.recv(1024)
##
##                ent.insert(END, output)
##                print output,
                channel.send('show arp no-resolve interface ' + Int.get() + '\n')
                time.sleep(3)
                output = channel.recv(1024)
                print output,
                mac = mac_pl(output)
                ent.insert(END, 'MAC %s\n' % mac)
                channel.send('quit\n')
                time.sleep(1)
                output = channel.recv(1024)
                print output,

        channel.send('quit\n')
        time.sleep(1)
        output = channel.recv(1024)
        print output,

        ssh.close()


def create_window():
    global ent, ip1, ip2, ip3, ipv1, ipv2
    window = Toplevel()
    window.title(address);
    ent = ScrolledText(window, width = 40, height = 20)
    ent.pack(expand=Y, fill=BOTH)
    if IP1.get() != 'x':
        network1 = ipcalc.Network(IP1.get())
        ip1 = network1.host_first() + 1
    if IP2.get() != 'x':
        network2 = ipcalc.Network(IP2.get())
        ip2 = network2.host_first() + 1
    if IP3.get() != 'x':
        network3 = ipcalc.Network(IP3.get())
        ip3 = network3.host_first() + 1
    if IPV1.get() != 'x':
        networkv1 = ipcalc.Network(IPV1.get())
        ipv1 = networkv1.host_first()
    if IPV2.get() != 'x':
        networkv2 = ipcalc.Network(IPV2.get())
        ipv2 = networkv2.host_first()
    pinger()


root = Tk()
root.title("EDDS")
root.iconbitmap(r'favicon.ico')
opt = StringVar()

w = OptionMenu(root, opt, *OPTIONS, command = get)
w.configure(width = WIDTH, anchor = W)
w.pack(fill = BOTH)

IP1 = StringVar()
IP1.set('x')
IP2 = StringVar()
IP2.set('x')
IP3 = StringVar()
IP3.set('x')
IPV1 = StringVar()
IPV1.set('x')
IPV2 = StringVar()
IPV2.set('x')
Int = StringVar()
Int.set('x')


Label(textvariable = IP1).pack(side = LEFT)
Label(textvariable = IP2).pack(side = LEFT)
Label(textvariable = IP3).pack(side = LEFT)
Label(textvariable = IPV1).pack(side = LEFT)
Label(textvariable = IPV2).pack(side = LEFT)
Label(textvariable = Int).pack(side = LEFT)
Button(text = 'Go', command=create_window).pack(fill = BOTH)

root.mainloop()

