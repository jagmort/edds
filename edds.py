#-*-coding: utf-8-*-
#!/usr/bin/python

import xlrd
import glob
from Tkinter import *
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
    LOST = 50

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
    print '=' * 80 + "\n" + str(int(found)) + ': ' + data[int(found) - 1][XLS.SERV] + ' ' + data[int(found) - 1][XLS.ADDR] + "\n"

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
        
    MAC.set('x')
    
    LIP1.configure(bg = 'SystemWindow')
    LIP2.configure(bg = 'SystemWindow')
    LIP3.configure(bg = 'SystemWindow')
    LIPV1.configure(bg = 'SystemWindow')
    LIPV2.configure(bg = 'SystemWindow')
    LInt.configure(bg = 'SystemWindow')
    LMAC.configure(bg = 'SystemWindow')


def ping_pl(text):
    for line in text.splitlines():
         result = re.search(r'([\d\.]+)% packet loss', line)
         if result:
             break
    return result.group(1)
    
def mac_pl(text):
    mac = ''
    for line in text.splitlines():
         result = re.search(r'([0-9a-fA-F]{2}(?::[0-9a-fA-F]{2}){5})', line)
         if result:
            mac = result.group(1)
            break
    return mac

def int_pl(text):
    status = ''
    for line in text.splitlines():
         result = re.search(r'Physical link is ([A-z]+)', line)
         if result:
            status = result.group(1)
            break
    return status
    
def pinger():
    global ent, ip1, ip2, ip3, ipv1, ipv2
    
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
    if float(ip1_ping) > 0:
        LIP1.configure(bg = 'red')
    else:
        LIP1.configure(bg = 'green')           
    channel.send(XLS.PING % str(ip2))
    time.sleep(XLS.SLEEP)
    output = channel.recv(65536)
    print output,
    ip2_ping = ping_pl(output)
    if float(ip2_ping) > 0:
        LIP2.configure(bg = 'red')
    else:
        LIP2.configure(bg = 'green')           

    if IP3.get() != 'x':
        channel.send(XLS.PING % str(ip3))
        time.sleep(XLS.SLEEP)
        output = channel.recv(65536)
        print output,
        ip3_ping = ping_pl(output)
        if float(ip3_ping) > 0:
            LIP3.configure(bg = 'red')
        else:
            LIP3.configure(bg = 'green')           

    if IPV1.get() != 'x':
        channel.send(XLS.PING % str(ipv1))
        time.sleep(XLS.SLEEP)
        output = channel.recv(65536)
        print output,
        ipv1_ping = ping_pl(output)
        if float(ipv1_ping) > 0:
            LIPV1.configure(bg = 'red')
        else:
            LIPV1.configure(bg = 'green')           

    if IPV2.get() != 'x':
        channel.send(XLS.PING % str(ipv2))
        time.sleep(XLS.SLEEP)
        output = channel.recv(65536)
        print output,
        ipv2_ping = ping_pl(output)
        if float(ipv2_ping) > 0:
            LIPV2.configure(bg = 'red')
        else:
            LIPV2.configure(bg = 'green')           

    if Int.get() != 'x':
        ip = ''
        if float(ip1_ping) < XLS.LOST:
            ip = str(ip1)
        elif float(ip2_ping) < XLS.LOST:
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

            channel.send('show interfaces ' + Int.get() + '\n')
            time.sleep(2)
            output = channel.recv(1024)
            print output,
            status = int_pl(output)
            if status != 'Up':
                LInt.configure(bg = 'red')
            else:
                LInt.configure(bg = 'green')           
                channel.send(' \n')
                time.sleep(2)
                output = channel.recv(1024)
                print output,

                channel.send('show arp no-resolve interface ' + Int.get() + '\n')
                time.sleep(2)
                output = channel.recv(1024)
                print output,
                mac = mac_pl(output)
                if mac != '':
                    MAC.set(mac)
                    LMAC.configure(bg = 'green')
                else:
                    LMAC.configure(bg = 'red')
                
            channel.send('quit\n')
            time.sleep(1)
            output = channel.recv(1024)
            print output,

    channel.send('quit\n')
    time.sleep(1)
    output = channel.recv(1024)
    print output

    ssh.close()

def simple():
    ip = ipcalc.IP(IPS.get())
    print '=' * 80 + "\nping " + str(ip) + '\n'
    
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
    channel.send(XLS.PING % str(ip))
    time.sleep(XLS.SLEEP)
    output = channel.recv(65536)
    print output,

    channel.send('quit\n')
    time.sleep(1)
    output = channel.recv(1024)
    print output

    ssh.close()


root = Tk()
root.title("EDDS")
root.iconbitmap(r'favicon.ico')
opt = StringVar()

fm00 = Frame(root, padx = 5)
Label(fm00, text = 'Адрес точки ЕДДС').pack(side = LEFT)
fm00.pack(fill = X, expand=YES)

fm0 = Frame(root, padx = 3)
w = OptionMenu(fm0, opt, *OPTIONS, command = get)
w.configure(width = WIDTH, anchor = W)
w.pack(fill = X)
fm0.pack(fill = X, expand=YES)

fm = Frame(root, padx = 5)
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
MAC = StringVar()
MAC.set('x')
IPS = StringVar()
IPS.set('10.220.0.1')

fmIP1 = Frame(fm)
Label(fmIP1, text = 'Основной').pack(fill = X)
LIP1 = Entry(fmIP1, textvariable = IP1, width = 16)
LIP1.pack(fill = X)
fmIP1.pack(side = LEFT)

fmIP2 = Frame(fm)
Label(fmIP2, text = 'Резервный').pack(fill = X)
LIP2 = Entry(fmIP2, textvariable = IP2, width = 16)
LIP2.pack(fill = X)
fmIP2.pack(side = LEFT)

fmIP3 = Frame(fm)
Label(fmIP3, text = '').pack(fill = X)
LIP3 = Entry(fmIP3, textvariable = IP3, width = 16)
LIP3.pack(fill = X)
fmIP3.pack(side = LEFT)

fmIPV1 = Frame(fm)
Label(fmIPV1, text = 'VipNet').pack(fill = X)
LIPV1 = Entry(fmIPV1, textvariable = IPV1, width = 16)
LIPV1.pack(fill = X)
fmIPV1.pack(side = LEFT)

fmIPV2 = Frame(fm)
Label(fmIPV2, text = '').pack(fill = X)
LIPV2 = Entry(fmIPV2, textvariable = IPV2, width = 16)
LIPV2.pack(fill = X)
fmIPV2.pack(side = LEFT)

fmInt = Frame(fm)
Label(fmInt, text = 'Порт на об. доступа').pack(fill = X)
LInt = Entry(fmInt, textvariable = Int, width = 10)
LInt.pack(fill = X)
fmInt.pack(side = LEFT)

fmMAC = Frame(fm)
Label(fmMAC, text = 'MAC конечного об.').pack(fill = X)
LMAC = Entry(fmMAC, textvariable = MAC, width = 16)
LMAC.pack(fill = X)
fmMAC.pack(side = LEFT)

fmGo = Frame(fm)
Label(fmGo, text = '').pack(fill = X)
Button(fmGo, text = 'Go', command=pinger, width = 6).pack(fill = X)
fmGo.pack(side = LEFT)
fm.pack(side = LEFT)

fm2 = Frame(root, padx = 5, pady = 3)
Label(fm2, text = '').pack(fill = BOTH)
Entry(fm2, textvariable = IPS, width = 14).pack(side = LEFT)
Button(fm2, text = 'Ping', command=simple, width = 6).pack(side = LEFT)
fm2.pack(fill = X, expand = YES)

root.mainloop()

