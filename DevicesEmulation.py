
import socket
import threading


from tkinter import *

from tkinter.ttk import Checkbutton




def client_delete(clients, i):
    try:
        clients[i].close()
    except OSError:
        pass
    clients.pop(i)


def clients_add_del(clients, ch_client_var):
    _client_count = int(en_client_count.get())
    if ch_client_var.get():
        if len(clients) == 0:
            for i in range(_client_count):
                _client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                try:
                    _client.connect((en_serv_ip.get(), int(en_serv_port.get())))
                except socket.timeout or OSError:
                    pass
                else:
                    _client.setblocking(0)
                    # print(f'{i}  ::  {_client}')
                    clients.append(_client)
    else:
        if len(clients) != 0:
            for i in reversed(range(len(clients))):
                try:
                    clients[i].close()
                except OSError:
                    pass
                else:
                    clients.pop(i)


def client_alive_check(clients, channel):

    for i in reversed(range(len(clients))):
        try:
            _data = clients[i].recv(4096)
        except socket.timeout or OSError:
            pass
        except OSError:
            pass
        else:
            if _data != b'':
                global packet_count
                _data = f'#CA01PM10{str(i).zfill(4)}\n#cmd#channelinfo#OK' \
                        f'#InterfaceType="{channel}"#Packet={packet_count}\n'
                packet_count += 1
                try:
                    clients[i].send(bytes(_data.encode('windows-1251')))
                except OSError:
                    pass
            else:
                client_delete(clients, i)

        global period_cur
        if period_cur == 0:
            _data = f'#CA01PM100{str(i).zfill(3)}\n#cmd#custopacket#OK' \
                    f'#gsm={len(ClientGSM)}#wifi={len(ClientWIFI)}#eth={len(ClientETH)}"#Packet={packet_count}\n'
            packet_count += 1
            try:
                clients[i].send(bytes(_data.encode('windows-1251')))
            except OSError:
                pass


def tcp_check():
    clients_add_del(ClientGSM, ch_client_gsm_var)
    clients_add_del(ClientETH, ch_client_eth_var)
    clients_add_del(ClientWIFI, ch_client_wifi_var)
    client_alive_check(ClientGSM, 'gsm')
    client_alive_check(ClientETH, 'ethernet')
    client_alive_check(ClientWIFI, 'wifi')

    global period_cur
    if ch_send_period_var.get() and period_cur != 0:
        period_cur -= 1
    else:
        period_cur = int(en_send_period.get()) * 2

    threading.Timer(0.5, tcp_check).start()



if __name__ == '__main__':
    serv_ip = '127.0.0.1'

    serv_port = 6665
    packet_count = 0

    ClientGSM = []
    ClientETH = []
    ClientWIFI = []

    window = Tk()
    window.title('DebugTest')
    window.geometry('650x50+500+100')

    lbl = Label(window, text='  ')
    lbl.pack(side=LEFT)

    en_serv_ip = Entry(window, width=15)
    en_serv_ip.pack(side=LEFT, padx=3)
    en_serv_ip.insert(0, serv_ip)

    lbl = Label(window, text=':')
    lbl.pack(side=LEFT)

    en_serv_port = Entry(window, width=5)
    en_serv_port.pack(side=LEFT, padx=3)
    en_serv_port.insert(0, serv_port)

    lbl = Label(window, text=' ClientCount =')
    lbl.pack(side=LEFT)

    en_client_count = Entry(window, width=4)
    en_client_count.pack(side=LEFT, padx=0)
    en_client_count.insert(0, 500)

    lbl = Label(window, text='    ')
    lbl.pack(side=LEFT)

    ch_send_period_var = BooleanVar()
    ch_send_period = Checkbutton(window, variable=ch_send_period_var)
    ch_send_period.pack(side=LEFT, padx=0)

    lbl = Label(window, text='SendPeriod =')
    lbl.pack(side=LEFT)

    en_send_period = Entry(window, width=4)
    en_send_period.pack(side=LEFT, padx=0)
    en_send_period.insert(0, 1)
    period_cur = int(en_send_period.get()) * 2

    lbl = Label(window, text='    ')
    lbl.pack(side=LEFT)

    ch_client_gsm_var = BooleanVar()
    ch_client_gsm = Checkbutton(window, variable=ch_client_gsm_var, text='gsm')
    ch_client_gsm.pack(side=LEFT, padx=12)

    ch_client_wifi_var = BooleanVar()
    ch_client_wifi = Checkbutton(window, variable=ch_client_wifi_var, text='wifi')
    ch_client_wifi.pack(side=LEFT, padx=12)

    ch_client_eth_var = BooleanVar()
    ch_client_eth = Checkbutton(window, variable=ch_client_eth_var, text='eth')
    ch_client_eth.pack(side=LEFT, padx=12)

    tcp_check()
    window.mainloop()
