import threading
from collections import deque
import datetime
import time
import winreg
from serial import win32 as win32
import ctypes  # Распределить по классам импорты

from tkinter import *
from tkinter.ttk import Combobox
from tkinter.ttk import Checkbutton
from tkinter.scrolledtext import *
import re
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import numpy as np
import os

class SerialException(IOError):  # Исключения в класс serialport
    """ """


class PortNotOpenError(SerialException):
    def __init__(self):
        super(PortNotOpenError, self).__init__('Attempting to use a port that is not open')


class SerialPort:
    QUEUE_SIZE = 4096

    def __init__(self, name, **kwargs):
        self.name = name
        self.baudrate = kwargs.pop('baudrate', 9600)
        self.format = kwargs.pop('format', '8n1')
        self._handle = None
        self._overlapped_rx = None
        self._overlapped_tx = None
        if kwargs:
            raise SerialException(f'Unknown option -{next(iter(kwargs.keys()))}')

    def configure(self, forced=False, **kwargs):
        baudrate = self.baudrate
        format_ = self.format
        self.baudrate = kwargs.pop('baudrate', self.baudrate)
        self.format = kwargs.pop('format', self.format)
        if kwargs:
            raise SerialException(f'Unknown option -{next(iter(kwargs.keys()))}')

        if self._handle is not None and (forced or baudrate != self.baudrate or format_ != self.format):
            fparity = 0
            parity = 0
            if self.format[1] == 'o':
                fparity = 1
                parity = win32.ODDPARITY
            elif self.format[1] == 'e':
                fparity = 1
                parity = win32.EVENPARITY

            if self.format[2] == '2':
                stopbits = win32.TWOSTOPBITS
            else:
                stopbits = win32.ONESTOPBIT

            try:
                bytesize = int(self.format[0])
            except ValueError:
                bytesize = 8

            dcb = win32.DCB(28, self.baudrate, 1, fparity, 0, 0, win32.DTR_CONTROL_DISABLE, 0, 0, 0, 0, 0, 0,
                            win32.RTS_CONTROL_DISABLE, 0, 0, 0, 0, 0, bytesize, parity, stopbits, 0, 0, 0, 0, 0, 0)

            timeouts = win32.COMMTIMEOUTS(ReadIntervalTimeout=win32.MAXDWORD)

            if not win32.SetCommTimeouts(self._handle, ctypes.byref(timeouts)):
                raise SerialException(f'Failed configure {self.name} (SetCommTimeouts): {ctypes.WinError()}')

            if not win32.SetCommState(self._handle, ctypes.byref(dcb)):
                raise SerialException(f'Failed configure {self.name} (SetCommState): {ctypes.WinError()}')

    def open(self, configure=True, **kwargs):
        if self._handle is not None:
            return
        self._handle = win32.CreateFile('\\\\.\\' + self.name, win32.GENERIC_READ | win32.GENERIC_WRITE,
                                        0, None, win32.OPEN_EXISTING, win32.FILE_FLAG_OVERLAPPED, 0)
        if self._handle == win32.INVALID_HANDLE_VALUE:
            self._handle = None
            raise SerialException(f'Failed open {self.name} (CreateFile): {ctypes.WinError()}')
        if configure:
            win32.SetupComm(self._handle, self.QUEUE_SIZE, self.QUEUE_SIZE)
            win32.PurgeComm(self._handle, win32.PURGE_TXCLEAR | win32.PURGE_TXABORT |
                            win32.PURGE_RXCLEAR | win32.PURGE_RXABORT)
            self._overlapped_rx = win32.OVERLAPPED()
            self._overlapped_tx = win32.OVERLAPPED()
            self.configure(forced=True, **kwargs)

    def close(self):
        if self._handle is not None:
            win32.CloseHandle(self._handle)
        self._handle = None

    def send(self, data, **kwargs):
        if self._handle is None:
            raise PortNotOpenError()
        if kwargs:
            self.configure(**kwargs)
        win32.WriteFile(self._handle, data, len(data), None, ctypes.byref(self._overlapped_tx))
        if win32.GetLastError() not in (win32.ERROR_SUCCESS, win32.ERROR_IO_PENDING):
            raise SerialException(f'Failed send {self.name} (WriteFile): {ctypes.WinError()}')

    def recv(self, size=None):
        if self._handle is None:
            raise PortNotOpenError()
        if size is None:
            size = self.status()[0]
        data = ctypes.create_string_buffer(size)
        win32.ReadFile(self._handle, data, size, None, ctypes.byref(self._overlapped_rx))
        if win32.GetLastError() not in (win32.ERROR_SUCCESS, win32.ERROR_IO_PENDING):
            raise SerialException(f'Failed receive {self.name} (ReadFile): {ctypes.WinError()}')
        return data.raw[:size]

    def status(self):
        if self._handle is None:
            raise PortNotOpenError()
        flags = win32.DWORD()
        comstat = win32.COMSTAT()
        if not (win32.ClearCommError(self._handle, ctypes.byref(flags), ctypes.byref(comstat))):
            raise SerialException(f'{self.name} (ClearCommError): {ctypes.WinError()}')
        # print(f'{comstat.cbInQue}/{comstat.cbOutQue}   {flags.value}')
        return comstat.cbInQue, comstat.cbOutQue

    def __del__(self):
        self.close()

    def __repr__(self):
        return f'{self.__class__.__name__}(id={id(self)}), {self.name}, {self.baudrate}, ' \
               f'{self.format}, handle={self._handle}'


class SerialPorts:
    PERIOD_HANDLER_RX_TX = 0.05  # seconds
    PERIOD_HANDLER_PORTS = 0.2  # seconds
    BUFFER_TX_SIZE = 10  # packets

    def __init__(self):
        self.ports = []
        self.thread_handler_ports = threading.Thread(target=self.handler_ports, daemon=True)
        self.thread_handler_ports.start()
        self.thread_handler_rx_tx = threading.Thread(target=self.handler_rx_tx, daemon=True)
        self.thread_handler_rx_tx.start()

    def __del__(self):
        self.thread_handler_ports.join()
        self.thread_handler_rx_tx.join()

    class _Port(SerialPort):

        def __init__(self, name, **kwargs):
            self.state = 'none'  # opened, closed, closing, busy, none
            self.to_open = False
            self.to_close = False
            self.timeout_interbytes = 1                             # NotUsed
            self.buffer_rx = kwargs.pop('buffer_rx', deque())
            self.timeout = kwargs.pop('timeout', 1)
            self.callback = kwargs.pop('callback', None)
            self.callback_thread = None
            self.buffer_tx = deque()
            self.que_rx_ = 0
            self.timeout_ = 0
            super(SerialPorts._Port, self).__init__(name, **kwargs)

        def configure(self, *args, **kwargs):
            self.buffer_rx = kwargs.pop('buffer_rx', self.buffer_rx)
            self.timeout = kwargs.pop('timeout', self.timeout)
            self.callback = kwargs.pop('callback', self.callback)
            if kwargs:
                super(SerialPorts._Port, self).configure(*args, **kwargs)

    def status(self, name):
        pass

    def _port_find(self, name):
        for port in self.ports:
            if port.name == name:
                return port

    def open(self, port, isname=True, **kwargs):
        if isname:
            port = self._port_find(port)
        if port is not None:
            if kwargs:
                self.configure(port, False, **kwargs)
            port.to_open = True
            return True
        return False

    def close(self, port, isname=True):
        if isname:
            port = self._port_find(port)
        if port is not None:
            port.to_close = True
            return True
        return False

    def configure(self, port, isname=True, **kwargs):
        if isname:
            port = self._port_find(port)
        if port is not None:
            if 'timeout' in kwargs:
                timeout = int(kwargs['timeout'] / self.PERIOD_HANDLER_RX_TX)
                if timeout < 1:
                    timeout = 1
                kwargs.update({'timeout': timeout})
            try:
                port.configure(**kwargs)
            except SerialException:
                pass
            return True
        return False

    def send(self, port, data, **kwargs):
        port = self._port_find(port)
        if port is not None:
            if port.state == 'opened':
                if len(port.buffer_tx) < self.BUFFER_TX_SIZE:
                    buffer_rx = kwargs.pop('buffer_rx', deque())
                    timeout = kwargs.pop('timeout', self.PERIOD_HANDLER_RX_TX)
                    callback = kwargs.pop('callback', None)
                    kwargs.update({'buffer_rx': buffer_rx, 'timeout': timeout, 'callback': callback})
                    port.buffer_tx.append((data, kwargs))
                    return True
            else:
                port.to_open = True
        return False

    def _close(self, port):
        port.state = 'closing'
        self.callback_ex(port)
        while port.buffer_tx:
            data, kw = port.buffer_tx.popleft()
            port.callback = kw.pop('callback', None)
            self.callback_ex(port)
        port.callback = None
        port.close()

    def handler_ports(self):
        while True:
            for count in range(5):
                ports_actual = []
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'HARDWARE\\DEVICEMAP\\SERIALCOMM')
                for i in range(256):
                    try:
                        ports_actual.append(winreg.EnumValue(key, i)[1])
                    except WindowsError:
                        break

                # Handling none ports
                for port in self.ports:
                    port_find = False
                    for name in ports_actual:
                        if port.name == name:
                            port_find = True
                            break
                    if not port_find:
                        if port.state == 'opened':
                            self._close(port)
                        port.state = 'none'

                # Adding new ports
                for name in ports_actual:
                    port_find = False
                    for port in self.ports:
                        if port.name == name:
                            if port.state == 'none':
                                port.state = 'closed'
                            port_find = True
                            break
                    if not port_find:
                        self.ports.append(self._Port(name))

                # Opening/Closing
                for port in self.ports:
                    if port.to_close is True and port.state == 'opened':
                        self._close(port)
                        port.state = 'closed'

                    elif port.to_open is True and port.state == 'closed':
                        try:
                            port.open()
                        except SerialException:
                            pass
                        else:
                            port.buffer_tx = deque()
                            port.buffer_rx.clear()
                            port.que_rx_ = 0
                            port.timeout_ = 0
                            port.state = 'opened'

                    port.to_close = False
                    port.to_open = False

                if count == 0:
                    # Scaning ports
                    for port in self.ports:
                        if port.state not in ('none', 'opened'):
                            a = time.time()
                            try:
                                port.open(False)
                            except SerialException:
                                port.state = 'busy'
                            else:
                                port.close()
                                port.state = 'closed'

                            a = time.time() - a
                            if a > 0.2:
                                print('\n' + f'{port.name} scanTime={a}')

                    # print('----------------------------------')
                    # for port in self.ports:
                    #     print(f'{port.name} {port.state}')

                time.sleep(self.PERIOD_HANDLER_PORTS)

    @staticmethod
    def callback_ex(port):
        if port.callback is not None:
            port.callback_thread = threading.Thread(target=port.callback, daemon=True)
            port.callback_thread.start()

    #                                     autoclose serial if not use(rx or tx), смотреть уходят ли TX
    def handler_rx_tx(self):
        while True:
            # asd = time.time()
            for port in self.ports:
                if port.state == 'opened' and port.to_close is not True:
                    try:
                        que_rx, que_tx = port.status()
                    except SerialException:
                        port.to_close = True
                    else:
                        # Timeouts
                        if port.timeout_ != 0 and que_tx == 0:
                            port.timeout_ -= 1
                            if port.timeout_ == 0:
                                self.callback_ex(port)

                        # Receive
                        if que_rx != 0 and port.que_rx_ == que_rx:
                            port.timeout_ = 0
                            try:
                                data = port.recv(que_rx)
                            except SerialException:
                                port.to_close = True
                            else:
                                # print(f'{port.name}: {data}')
                                port.buffer_rx.append(data)
                                self.callback_ex(port)
                        else:
                            port.que_rx_ = que_rx

                        # Send (must be after Recieve because of timeout_)
                        if port.buffer_tx and que_tx == 0 and port.timeout_ == 0:
                            data, kw = port.buffer_tx.popleft()
                            self.configure(port, False, **kw)
                            try:
                                port.send(data)
                            except SerialException:
                                port.to_close = True
                            else:
                                port.timeout_ = port.timeout
            # print(time.time() - asd)
            time.sleep(self.PERIOD_HANDLER_RX_TX)


def event_files_scan(event):
    files = CsvFile.listdir()
    files.reverse()
    combo_files['values'] = files


def event_file_select(event):
    window.after(100, file_select)


def file_select():
    CsvFile.file = combo_files.get()
    file = CsvFile.read(CsvFile.file)
    line = file[0]
    data = re.split(r'[\s;]+', line)

    for i in reversed(range(len(data))):
        index = data[i].find('(')
        if data[i] == 'Time' or data[i] == '':
            data.pop(i)
        elif index >= 0:
            data[i] = data[i][:index]

    params_update(set(data))


class CsvFile:
    dir = os.path.join(os.getcwd(), 'Logs')
    file = ''
    file_ext = '.csv'

    @classmethod
    def write(cls, file, data):
        dir_ok = True
        try:
            os.mkdir(cls.dir)
        except FileExistsError:
            pass
        except OSError:
            dir_ok = False
        if dir_ok:
            try:
                file = open(os.path.join(cls.dir, file + cls.file_ext), 'w')
            except OSError:
                pass
            else:
                file.write(data)
                file.close()

    @classmethod
    def read(cls, file):
        try:
            data = open(os.path.join(cls.dir, file + cls.file_ext), 'r').readlines()
        except OSError:
            pass
        else:
            return data

    @classmethod
    def listdir(cls):
        files = []
        for file in os.listdir(cls.dir):
            index = file.find('.')
            if index >= 0:
                file = file[:index]
            files.append(file)
        return files


DustData = {}


def btn_graph_command():
    _thread = threading.Thread(target=log_show, daemon=True)
    _thread.start()
    btn_graph.configure(state='disabled')


def log_show(to_graph=True):
    global DustData
    if btn_task['text'] != 'Start':
        data = 'Time;'
        for serial in DustData['Serials'].keys():
            for param in DustData['Serials'][serial].keys():
                data += f'{param}({serial});'

        for index in range(len(DustData['Time'])):
            line = f"\n{DustData['Time'][index]};"
            for serial in DustData['Serials'].keys():
                for param in DustData['Serials'][serial].keys():
                    line += f"{DustData['Serials'][serial][param][index]};".replace('.', ',')
            data += line

        CsvFile.write(CsvFile.file, data)

    if to_graph:
        csv = {}
        file = CsvFile.read(CsvFile.file)
        line_first = True
        for line in file:
            data = re.split(r'[\s;]+', line)
            if line_first:
                line_first = False
                params = []
                for index, param in enumerate(DustData['Params']):
                    if DustData['Params2Graph'][index].get() is True:
                        params.append(param)
                for index, param in enumerate(data):
                    if param == 'Time':
                        csv.update({'Time': {'Index': index, 'Values': []}})
                    for _param in params:
                        if param.find(_param) >= 0:
                            csv.update({param: {'Index': index, 'Values': []}})
                            break
            else:
                for param in csv.keys():
                    csv[param]['Values'].append(data[csv[param]['Index']].replace(',', '.'))

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        # fig.update_xaxes(title_text="")
        fig.update_yaxes(title_text="", secondary_y=False)
        fig.update_yaxes(title_text="", secondary_y=True, showgrid=False)
        colors = ('blue', 'DeepSkyBlue', 'green', 'red', 'magenta',
                  'purple', 'GoldenRod', 'black', 'SaddleBrown', 'Teal')
        serial = ''
        index = 0
        for param in csv.keys():
            if param == 'Time':
                continue

            regular = re.search(r'\((\w+)\)', param)
            if regular is not None:
                if serial != regular.group(1):
                    index += 1
                serial = regular.group(1)
            else:
                index += 1
                serial = ''

            if index > len(colors) - 1:
                index = 0

            fig.add_trace(go.Scatter(x=csv['Time']['Values'], y=np.array(csv[param]['Values'], dtype=float),
                                     mode='lines', marker=dict(color=colors[index]), name=param),
                          secondary_y=False)

        fig.update_layout(legend_orientation="h", legend=dict(x=.5, xanchor="center"),
                          title='',
                          barmode='group', hovermode='x',
                          margin=dict(l=0, r=0, t=30, b=0))
        fig.update_traces(hoverinfo="x+y+name", hovertemplate='  %{y} ')
        btn_graph.configure(state='normal')
        fig.show()


def params_update(params):
    global DustData
    DustData['Params'] = params
    global chk_widgets
    for chk in chk_widgets:
        chk.destroy()
    DustData['Params2Graph'] = []
    chk_widgets = []
    for param in DustData['Params']:
        var = BooleanVar()
        chk = Checkbutton(frame_settings2, text=param + '   ', variable=var)
        chk.pack(side=LEFT, padx=0)
        var.set(True)
        chk_widgets.append(chk)
        DustData['Params2Graph'].append(var)


def task_toggle():
    global DustData
    if btn_task['text'] == 'Start':
        btn_task.configure(text="STOP", fg='red')
        en_params.configure(state='disable')
        combo_files.set('')
        combo_files.configure(state='disable')
        text_log_clear()
        CsvFile.file = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        DustData.clear()
        DustData = {'Serials': {},
                    'Time': [],
                    'Params': [],
                    'Params2Graph': []
                    }
        params = re.split(r'[\s;,\.]+', en_params.get())
        if params[len(params) - 1] == '':
            params.pop(len(params) - 1)
        params_update(params)
        window.after(50, callback)
    else:
        log_show(False)
        params_update([])
        btn_task.configure(text="Start", fg='black')
        en_params.configure(state='normal')
        combo_files.configure(state='readonly')
        text_log.configure(state='normal')
        text_log.delete(1.0, END)
        text_log.configure(state='disabled')
        text_log.yview(1.0)


def datalog1_menu(event):
    text_log_menu.post(text_log.winfo_rootx() + event.x, text_log.winfo_rooty() + event.y)


def text_log_clear(change_state=True):
    if change_state:
        text_log.configure(state='normal')
    text_log.delete(1.0, END)
    text_log.yview(END)
    if change_state:
        text_log.configure(state='disabled')


def callback():
    if btn_task['text'] == 'Start':
        return
    for port in Serial_Ports.ports:
        packets = re.split(r'[\s;]+', en_send.get())
        for packet in packets:
            send = packet.encode('windows-1251') + b'\n\r'
            Serial_Ports.send(port.name, send, baudrate=115200, format='8n1', timeout=0,
                              callback=None, buffer_rx=Buf3)

    while Buf3:
        data = Buf3.popleft()
        data = data.decode('windows-1251')
        reg = re.match(r'ID=([0-9ABCDEF]+).+\s', data)
        if reg is None:
            continue
        serial = reg.group(1)

        params_val = []
        for param in DustData['Params']:
            reg = re.search(r'{0}=([0-9\.]+)\s'.format(param), data)
            value = 0.0
            if reg is not None:
                try:
                    value = float(reg.group(1))
                except ValueError:
                    pass
            params_val.append(value)

        device_find = False
        for _serial in DustData['Serials'].keys():
            if _serial == serial:
                device_find = True
                break

        if not device_find:
            DustData['Serials'].update({serial: {}})
            for param in DustData['Params']:
                DustData['Serials'][serial].update({param: []})
                for i in range(len(DustData['Time'])):
                    DustData['Serials'][serial][param].append(0.0)

        for index, param in enumerate(DustData['Params']):
            DustData['Serials'][serial][param].append(params_val[index])

    time_now = datetime.datetime.now()
    DustData['Time'].append(time_now.strftime('%H:%M:%S'))

    for serial in DustData['Serials'].keys():
        for param in DustData['Serials'][serial].keys():
            if len(DustData['Serials'][serial][param]) != len(DustData['Time']):
                DustData['Serials'][serial][param].append(0.0)

    log = 'Time='
    index = len(DustData['Time']) - 1
    log += DustData['Time'][index]
    for serial in DustData['Serials'].keys():
        log += '\n' + serial.ljust(20)
        for key, val in DustData['Serials'][serial].items():
            log += f'{key}={int(val[index])}'.ljust(12)
    text_log.configure(state='normal')
    text_log.delete(1.0, END)
    text_log.insert(END, log + '\n\n')
    text_log.configure(state='disabled')
    text_log.yview(1.0)

    _period = en_period.get()
    try:
        period = int(_period)
    except ValueError:
        period = 2
    if period < 2:
        period = 2
    if period > 10:
        period = 10
    if period != _period:
        en_period.delete(0, END)
        en_period.insert(END, period)
    period *= 1000
    window.after(period, callback)


if __name__ == '__main__':
    window = Tk()
    window.period_check_id = 1
    window.title('DustMeasure')
    window.geometry('900x400+200+20')
    frame_settings = Frame(window, width=30)
    frame_settings.pack(side=TOP, anchor=NW, padx=6, pady=6, fill=X)

    btn_task = Button(frame_settings, width=10, height=1, text="Start", command=task_toggle)
    btn_task.pack(side=LEFT, padx=3)

    lbl = Label(frame_settings, text='    Send= ')
    lbl.pack(side=LEFT)

    en_send = Entry(frame_settings, width=20)
    en_send.insert(END, 'getall')
    en_send.pack(side=LEFT)

    lbl = Label(frame_settings, text='    Period(sec)= ')
    lbl.pack(side=LEFT)

    en_period = Entry(frame_settings, width=3)
    en_period.insert(END, '2')
    en_period.pack(side=LEFT)

    lbl = Label(frame_settings, text='    Params= ')
    lbl.pack(side=LEFT)

    en_params = Entry(frame_settings, width=40)
    en_params.insert(END, 'PM2 PR2 PM9 PR9 TSPM TSPR')
    en_params.pack(side=LEFT)

    frame_settings2 = Frame(window, width=30)
    frame_settings2.pack(side=TOP, anchor=NW, padx=6, pady=6, fill=X)

    btn_graph = Button(frame_settings2, width=10, height=1, text="Graph", command=btn_graph_command)
    btn_graph.pack(side=LEFT, padx=3)

    lbl = Label(frame_settings2, text='    OpenFile=')
    lbl.pack(side=LEFT)

    combo_files = Combobox(frame_settings2, state='readonly', width=20)
    combo_files.pack(side=LEFT, padx=3)
    combo_files.bind('<Button-1>', event_files_scan)
    combo_files.bind('<<ComboboxSelected>>', event_file_select)

    lbl = Label(frame_settings2, text='    ')
    lbl.pack(side=LEFT)

    frame_log = Frame(window, width=300)
    frame_log.pack(side=TOP, anchor=NW, expand=1, fill=BOTH)

    chk_vars = []
    chk_widgets = []

    text_log = ScrolledText(master=frame_log, height=40, wrap='word')
    text_log.configure(state='disabled')
    text_log.bind("<Button-3>", datalog1_menu)
    text_log.pack(side=TOP, fill=BOTH, expand=1, anchor=NW, padx=6)

    text_log_menu = Menu(frame_log, tearoff=0)
    text_log_menu.add_command(label="Clear", command=text_log_clear)

    Serial_Ports = SerialPorts()
    Buf3 = deque()
    window.mainloop()
