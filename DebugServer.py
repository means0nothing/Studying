import serial
import sys
import socket
from ftplib import FTP
import threading
import os
import copy
import winreg
import datetime
import time

from tkinter import *
from tkinter.ttk import Style
from tkinter.ttk import Combobox
from tkinter.ttk import Checkbutton
from tkinter.ttk import Scrollbar
from tkinter.ttk import Notebook
from tkinter.ttk import Treeview
from tkinter.scrolledtext import *
from tkinter.font import *

device_lists = {}
client_buffer = []
tree_chosen = ''
tree_selected = ''
tree_expand_all = False
tree_collapse_all = False
tx_queue_buffer = []
DATALOG_BUFFER = 100
timing_en = False
count_3d2 = 0
count_mppt = 0
crc_tab = [0x0000, 0x1021, 0x2042, 0x3063, 0x4084, 0x50A5, 0x60C6, 0x70E7,
           0x8108, 0x9129, 0xA14A, 0xB16B, 0xC18C, 0xD1AD, 0xE1CE, 0xF1EF,
           0x1231, 0x0210, 0x3273, 0x2252, 0x52B5, 0x4294, 0x72F7, 0x62D6,
           0x9339, 0x8318, 0xB37B, 0xA35A, 0xD3BD, 0xC39C, 0xF3FF, 0xE3DE,
           0x2462, 0x3443, 0x0420, 0x1401, 0x64E6, 0x74C7, 0x44A4, 0x5485,
           0xA56A, 0xB54B, 0x8528, 0x9509, 0xE5EE, 0xF5CF, 0xC5AC, 0xD58D,
           0x3653, 0x2672, 0x1611, 0x0630, 0x76D7, 0x66F6, 0x5695, 0x46B4,
           0xB75B, 0xA77A, 0x9719, 0x8738, 0xF7DF, 0xE7FE, 0xD79D, 0xC7BC,
           0x48C4, 0x58E5, 0x6886, 0x78A7, 0x0840, 0x1861, 0x2802, 0x3823,
           0xC9CC, 0xD9ED, 0xE98E, 0xF9AF, 0x8948, 0x9969, 0xA90A, 0xB92B,
           0x5AF5, 0x4AD4, 0x7AB7, 0x6A96, 0x1A71, 0x0A50, 0x3A33, 0x2A12,
           0xDBFD, 0xCBDC, 0xFBBF, 0xEB9E, 0x9B79, 0x8B58, 0xBB3B, 0xAB1A,
           0x6CA6, 0x7C87, 0x4CE4, 0x5CC5, 0x2C22, 0x3C03, 0x0C60, 0x1C41,
           0xEDAE, 0xFD8F, 0xCDEC, 0xDDCD, 0xAD2A, 0xBD0B, 0x8D68, 0x9D49,
           0x7E97, 0x6EB6, 0x5ED5, 0x4EF4, 0x3E13, 0x2E32, 0x1E51, 0x0E70,
           0xFF9F, 0xEFBE, 0xDFDD, 0xCFFC, 0xBF1B, 0xAF3A, 0x9F59, 0x8F78,
           0x9188, 0x81A9, 0xB1CA, 0xA1EB, 0xD10C, 0xC12D, 0xF14E, 0xE16F,
           0x1080, 0x00A1, 0x30C2, 0x20E3, 0x5004, 0x4025, 0x7046, 0x6067,
           0x83B9, 0x9398, 0xA3FB, 0xB3DA, 0xC33D, 0xD31C, 0xE37F, 0xF35E,
           0x02B1, 0x1290, 0x22F3, 0x32D2, 0x4235, 0x5214, 0x6277, 0x7256,
           0xB5EA, 0xA5CB, 0x95A8, 0x8589, 0xF56E, 0xE54F, 0xD52C, 0xC50D,
           0x34E2, 0x24C3, 0x14A0, 0x0481, 0x7466, 0x6447, 0x5424, 0x4405,
           0xA7DB, 0xB7FA, 0x8799, 0x97B8, 0xE75F, 0xF77E, 0xC71D, 0xD73C,
           0x26D3, 0x36F2, 0x0691, 0x16B0, 0x6657, 0x7676, 0x4615, 0x5634,
           0xD94C, 0xC96D, 0xF90E, 0xE92F, 0x99C8, 0x89E9, 0xB98A, 0xA9AB,
           0x5844, 0x4865, 0x7806, 0x6827, 0x18C0, 0x08E1, 0x3882, 0x28A3,
           0xCB7D, 0xDB5C, 0xEB3F, 0xFB1E, 0x8BF9, 0x9BD8, 0xABBB, 0xBB9A,
           0x4A75, 0x5A54, 0x6A37, 0x7A16, 0x0AF1, 0x1AD0, 0x2AB3, 0x3A92,
           0xFD2E, 0xED0F, 0xDD6C, 0xCD4D, 0xBDAA, 0xAD8B, 0x9DE8, 0x8DC9,
           0x7C26, 0x6C07, 0x5C64, 0x4C45, 0x3CA2, 0x2C83, 0x1CE0, 0x0CC1,
           0xEF1F, 0xFF3E, 0xCF5D, 0xDF7C, 0xAF9B, 0xBFBA, 0x8FD9, 0x9FF8,
           0x6E17, 0x7E36, 0x4E55, 0x5E74, 0x2E93, 0x3EB2, 0x0ED1, 0x1EF0]

CmdExample = {
    'server': {
        'EnabledServers': {'Value': '', 'Type': 'Label'},
        'Index': {'Value': '', 'Type': 'Combo', 'Choice': '0/1'},
        'Enabled': {'Value': '', 'Type': 'Combo', 'Choice': 'yes/no'},
        'Connected': {'Value': '', 'Type': 'Label'},
        'Channel': {'Value': '', 'Type': 'Combo', 'Choice': 'GSM/Ethernet/WiFi'},
        'Url': {'Value': '', 'Type': 'Combo', 'Choice': 'upload.cityair.ru/cityair.uniscan.biz'},
        'Port': {'Value': '', 'Type': 'Combo', 'Choice': '49041/6655'},
        'AesEnabled': {'Value': '', 'Type': 'Combo', 'Choice': 'yes/no'},
        'AesKey': {'Value': '', 'Type': 'Combo'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'datadelivery': {
        'Index': {'Value': '', 'Type': 'Label'},
        'LastDeliveredRecordNumber': {'Value': '', 'Type': 'Label'},
        'MostFreshDeliveredRecord': {'Value': '', 'Type': 'Label'},
        'FirstNotDelivered': {'Value': '', 'Type': 'Label'},
        'Period': {'Value': '', 'Type': 'Entry'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'get': {
        'Uptime': {'Value': '', 'Type': 'Label'},
        'DateTime': {'Value': '', 'Type': 'Label'},
        'GNSS': {'Value': '', 'Type': 'Label'},
        'GnssGpsCount': {'Value': '', 'Type': 'Label'},
        'GnssGlonassCount': {'Value': '', 'Type': 'Label'},
        'GnssLat': {'Value': '', 'Type': 'Label'},
        'GnssLon': {'Value': '', 'Type': 'Label'},
        'GnssDateTime': {'Value': '', 'Type': 'Label'},
        'T': {'Value': '', 'Type': 'Label'},
        'H': {'Value': '', 'Type': 'Label'},
        'P': {'Value': '', 'Type': 'Label'},
        'Dust2.5': {'Value': '', 'Type': 'Label'},
        'Dust9': {'Value': '', 'Type': 'Label'},
        'DustTsp': {'Value': '', 'Type': 'Label'},
        'DustBoxId': {'Value': '', 'Type': 'Label'},
        'DustBoxVersion': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
    },
    'time': {
        'DateTime': {'Value': '', 'Type': 'Label'},
        'Sync': {'Value': '', 'Type': 'Combo', 'Choice': 'GNSS/NTP'},
        'NtpChannel': {'Value': '', 'Type': 'Combo', 'Choice': 'GSM/Ethernet/WiFi'},
        'NtpUrl': {'Value': '', 'Type': 'Combo', 'Choice': 'upload.cityair.ru/cityair.uniscan.biz'},
        'NtpPort': {'Value': '', 'Type': 'Combo', 'Choice': '123/6655'},
        'NtpPeriod': {'Value': '', 'Type': 'Combo', 'Choice': '60'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'location': {
        'Mode': {'Value': '', 'Type': 'Combo', 'Choice': 'GNSS/Custom'},
        'Lat': {'Value': '', 'Type': 'Entry'},
        'Lon': {'Value': '', 'Type': 'Entry'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'debugserver': {
        'Connected': {'Value': '', 'Type': 'Label'},
        'Enabled': {'Value': '', 'Type': 'Combo', 'Choice': 'yes/no'},
        'Url': {'Value': '', 'Type': 'Combo', 'Choice': 'upload.cityair.ru/cityair.uniscan.biz'},
        'Port': {'Value': '', 'Type': 'Combo', 'Choice': '6655'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'gsm': {
        'Enabled': {'Value': '', 'Type': 'Label'},
        'State': {'Value': '', 'Type': 'Label'},
        'APN': {'Value': '', 'Type': 'Combo', 'Choice': 'internet'},
        'Login': {'Value': '', 'Type': 'Combo', 'Choice': 'gdata'},
        'Password': {'Value': '', 'Type': 'Combo', 'Choice': 'gdata'},
        'IMEI': {'Value': '', 'Type': 'Label'},
        'IMSI': {'Value': '', 'Type': 'Label'},
        'ICCID': {'Value': '', 'Type': 'Label'},
        'Rssi': {'Value': '', 'Type': 'Label'},
        'BitErrorRate': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'dust': {
        'Mode': {'Value': '', 'Type': 'Combo', 'Choice': 'Auto/Manual'},
        'Sensor': {'Value': '', 'Type': 'Combo', 'Choice': '1/2'},
        'ForceHeat': {'Value': '', 'Type': 'Combo', 'Choice': 'yes/no'},
        'Source': {'Value': '', 'Type': 'Label'},
        'Dust2.5': {'Value': '', 'Type': 'Label'},
        'Dust9': {'Value': '', 'Type': 'Label'},
        'DustWT': {'Value': '', 'Type': 'Label'},
        'DustV': {'Value': '', 'Type': 'Label'},
        'check_Fan': {'Value': '', 'Type': 'Label'},
        'HeaterCur': {'Value': '', 'Type': 'Label'},
        'DustT': {'Value': '', 'Type': 'Label'},
        'T0': {'Value': '', 'Type': 'Label'},
        'T1': {'Value': '', 'Type': 'Label'},
        'T2': {'Value': '', 'Type': 'Label'},
        'T3': {'Value': '', 'Type': 'Label'},
        'T4': {'Value': '', 'Type': 'Label'},
        'Sensor1Online': {'Value': '', 'Type': 'Label'},
        'Sensor1Time': {'Value': '', 'Type': 'Label'},
        'Sensor1MachineHours': {'Value': '', 'Type': 'Label'},
        'Sensor1_2.5': {'Value': '', 'Type': 'Label'},
        'Sensor1_9': {'Value': '', 'Type': 'Label'},
        'PMTempOut': {'Value': '', 'Type': 'Label'},
        'PMCur': {'Value': '', 'Type': 'Label'},
        'Sensor2Online': {'Value': '', 'Type': 'Label'},
        'Sensor2Time': {'Value': '', 'Type': 'Label'},
        'Sensor2MachineHours': {'Value': '', 'Type': 'Label'},
        'Sensor2_2.5': {'Value': '', 'Type': 'Label'},
        'Sensor2_9': {'Value': '', 'Type': 'Label'},
        'PRTempOut': {'Value': '', 'Type': 'Label'},
        'PRCur': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'wifi': {
        'Enabled': {'Value': '', 'Type': 'Label'},
        'SSID': {'Value': '', 'Type': 'Combo', 'Choice': 'Public'},
        'Password': {'Value': '', 'Type': 'Combo', 'Choice': 'neqwerty123'},
        'RssiValues': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'ethernet': {
        'Enabled': {'Value': '', 'Type': 'Label'},
        'State': {'Value': '', 'Type': 'Label'},
        'Mode': {'Value': '', 'Type': 'Combo', 'Choice': 'auto/manual'},
        'Mac': {'Value': '', 'Type': 'Entry'},
        'Ip': {'Value': '', 'Type': 'Entry'},
        'SubnetMask': {'Value': '', 'Type': 'Entry'},
        'Dns': {'Value': '', 'Type': 'Entry'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'interfaceboard': {
        'AutoFlash': {'Value': '', 'Type': 'Combo', 'Choice': 'on/off'},
        'BootState': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'temperature': {
        'Temperature': {'Value': '', 'Type': 'Label'},
        'TemperatureBeforeCorrection': {'Value': '', 'Type': 'Label'},
        'Intercept': {'Value': '', 'Type': 'Entry'},
        'Slope': {'Value': '', 'Type': 'Entry'},
        'Threshold': {'Value': '', 'Type': 'Entry'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'version': {
        'SoftwareVersion': {'Value': '', 'Type': 'Label'},
        'BootVersion': {'Value': '', 'Type': 'Label'},
        'InterfaceBoardVersion': {'Value': '', 'Type': 'Label'},
        'RebootCause': {'Value': '', 'Type': 'Label'},
        'RebootTimeStamp': {'Value': '', 'Type': 'Label'},
        'RebootCounter': {'Value': '', 'Type': 'Label'},
        'MD5': {'Value': '', 'Type': 'Label'},
        'Error': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'sensorGroup': {
        'Tph': {'Value': '', 'Type': 'Combo', 'Choice': '0'},
        'Dust': {'Value': '', 'Type': 'Combo', 'Choice': '30'},
        'Owen': {'Value': '', 'Type': 'Combo', 'Choice': '0'},
        'G1': {'Value': '', 'Type': 'Combo', 'Choice': '30'},
        'G2': {'Value': '', 'Type': 'Combo', 'Choice': '30'},
        'G3': {'Value': '', 'Type': 'Combo', 'Choice': '0'},
        'Wind': {'Value': '', 'Type': 'Combo', 'Choice': '0'},
        'Precipitation': {'Value': '', 'Type': 'Combo', 'Choice': '0'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'check': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'reboot': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'firmwareupdate': {
        'StartupError': {'Value': '', 'Type': 'Label'},
        'LoaderError': {'Value': '', 'Type': 'Label'},
        'DownloaderError': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
    },
    'channelinfo': {
        'InterfaceType': {'Value': '', 'Type': 'Label'},
        'ConnectionIndex': {'Value': '', 'Type': 'Label'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
    },
    'history': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'measure': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'modbus': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'requestemulator': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'modbusrequest': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'telemetry': {
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
    'gnss': {
        'enabled': {'Value': '', 'Type': 'Combo', 'Choice': 'yes/no'},
        'ButtonRead': {'Value': 'read', 'Type': 'Button'},
        'ButtonWrite': {'Value': 'write', 'Type': 'Button'},
    },
    'deviceconfig#default': {
        'Disable': {'Value': 'write', 'Type': 'Button'},
    },
}


def parse_ca(_string, _dict):
    for key, val in _dict.items():
        if _string.find('#' + key + '#') >= 0:
            for key11, val11 in _dict[key].items():
                if _string.find(key11 + '=') >= 0:
                    index = _string.find(key11 + '=') + len(key11 + '=')
                    if index >= len(_string):
                        continue
                    if _string[index] == '\"':
                        _dict[key][key11]['Value'] = _string[index + 1:_string.find('\"', index + 1)]
                    elif _string[index] == '{':
                        _dict[key][key11]['Value'] = _string[index + 1:_string.find('}', index + 1)]
                    else:
                        if _string.find(' ', index) >= 0:
                            _dict[key][key11]['Value'] = _string[index:_string.find(' ', index)]
                        else:
                            _dict[key][key11]['Value'] = _string[index:_string.find('\n', index)]
                else:
                    _dict[key][key11]['Value'] = ''


class EntryAutoFill:
    def __init__(self, parent, side, mode):
        self.parent = parent
        self.lbl_list = []
        self.lbl_list_active = 0
        self.auto_index = 0
        self.auto_var = []
        self.font = Font(family="TkDefaultFont", size=9, weight="normal")
        self.mode = mode
        self.win = Toplevel(parent)
        self.win.overrideredirect(True)
        self.win.withdraw()
        self.frame = Frame(self.win, highlightbackground="grey", highlightthickness=1)
        self.frame.pack()
        for i in range(30):
            lbl = Label(self.frame, font=self.font, bg='white', anchor="w")
            self.lbl_list.append(lbl)
        self.en = Entry(parent, name='en', width=600, validate="key", font=self.font)
        self.en['validatecommand'] = (self.en.register(self.en_val), '%P', '%d')
        self.en.pack(side=side, pady=6, padx=6, anchor=NW)
        self.en.bind('<Key>', self.en_key)
        self.en.bind('<Button-1>', self.en_key)

    def delete(self, first, last):
        return self.en.delete(first, last)

    def insert(self, index, string):
        return self.en.insert(index, string)

    def get(self):
        return self.en.get()

    def en_val(self, _string, _type):
        window.after(10, lambda x=_string: self.en_auto(x))
        return True

    def en_key(self, event):
        if event.keysym == 'Tab':
            if self.win.state() == 'normal':
                _string = self.en.get() + self.lbl_list[self.lbl_list_active]['text'][self.auto_index:]
                self.en.delete(0, END)
                self.en.insert(0, _string)
                return 'break'
        elif event.keysym == 'Return':
            if self.mode == 'group':
                en_command_group_event()
            elif self.mode == 'single':
                en_command_event()
        elif event.keysym == 'Down' or event.keysym == 'Up':
            if self.win.state() == 'normal':
                self.lbl_list[self.lbl_list_active]['bg'] = 'white'
                if event.keysym == 'Down':
                    if self.lbl_list_active == len(self.auto_var) - 1:
                        self.lbl_list_active = 0
                    else:
                        self.lbl_list_active += 1

                elif event.keysym == 'Up':
                    if self.lbl_list_active == 0:
                        self.lbl_list_active = len(self.auto_var) - 1
                    else:
                        self.lbl_list_active -= 1

                self.lbl_list[self.lbl_list_active]['bg'] = 'grey88'
        elif event.num == 1 or event.keysym == 'Left' or event.keysym == 'Right':
            window.after(100, lambda x=self.en.get(): self.en_auto(x))

    def en_auto(self, _string):
        if self.en.index(INSERT) != self.en.index(END):
            # print(f'{self.en.index(INSERT)} : {self.en.index(END)}')
            self.win.withdraw()
            return
        auto_var_width = 1
        param = ''
        self.auto_index = len(_string)
        index = 0
        self.auto_var = []
        for key in CmdExample.keys():
            keyl = key.lower()
            if _string.find(keyl) == 0:
                param = key
            elif keyl.find(_string) == 0:
                index = 0
                if len(keyl) > auto_var_width:
                    auto_var_width = len(keyl)
                self.auto_var.append(keyl)

        if param != '':
            for key, val in CmdExample[param].items():
                keyl = key.lower()
                index = _string.rfind('#') + 1
                if index < 0:
                    break
                _type = CmdExample[param][key]['Type']
                if _type == 'Combo' or _type == 'Entry':
                    if _string.find(keyl + '=') == index:
                        index += len(keyl) + 1
                        try:
                            _choice = CmdExample[param][key]['Choice']
                        except KeyError:
                            pass
                        else:
                            _choice = _choice.split('/')
                            for _item in _choice:
                                _item = _item.lower()
                                if _string[index:].find(_item) == 0:
                                    break
                                elif _item.find(_string[index:]) == 0:
                                    if len(_item) > auto_var_width:
                                        auto_var_width = len(_item)
                                    self.auto_var.append(_item)
                        break
                    elif _string.find(keyl) >= 0:
                        continue
                    elif keyl.find(_string[index:]) >= 0:
                        if len(keyl) + 1 > auto_var_width:
                            auto_var_width = len(keyl) + 1
                        self.auto_var.append(keyl + '=')
            self.auto_index -= index

        if len(self.auto_var) != 0:
            for i in range(len(self.lbl_list)):
                if i >= len(self.auto_var):
                    self.lbl_list[i].forget()
                    continue
                self.lbl_list[i]['width'] = auto_var_width
                self.lbl_list[i]['text'] = self.auto_var[i]
                self.lbl_list[i].pack(side=TOP)

            self.lbl_list[self.lbl_list_active]['bg'] = 'white'
            self.lbl_list_active = 0
            self.lbl_list[self.lbl_list_active]['bg'] = 'grey88'
            self.win.geometry(f'+{self.en.winfo_rootx() + self.font.measure(_string[:index])}'
                              f'+{self.en.winfo_rooty() + self.en.winfo_height()}')
            self.win.deiconify()
        else:
            self.win.withdraw()


class FwUpdateInfo:
    def __init__(self):
        self.channel = ''
        self.server = ''
        self.port = ''
        self.login = ''
        self.password = ''
        self.path = ''
        self.crc = ''
        self.path2 = ''
        self.crc2 = ''

    ftp_users_quantity = 21
    ftp_users_name = 'fw'
    ftp_users_pass = '\"\"'
    ftp_users_available = []

    @classmethod
    def ftp_users_init(cls):
        cls.ftp_users_available = []
        for i in range(cls.ftp_users_quantity):
            cls.ftp_users_available.append(True)

    @classmethod
    def ftp_users_release(cls, user):
        cls.ftp_users_available[int(user.replace(cls.ftp_users_name, ''))] = True

    @classmethod
    def ftp_users_rent(cls):
        for i in range(len(cls.ftp_users_available)):
            if cls.ftp_users_available[i]:
                cls.ftp_users_available[i] = False
                return f'{cls.ftp_users_name}{i}'
        return ''

    def sufficient(self):
        _info_sufficient = True
        if self.channel == '':
            _info_sufficient = False
        if self.server == '':
            _info_sufficient = False
        if self.port == '':
            _info_sufficient = False
        if self.login == '':
            _info_sufficient = False
        if self.password == '':
            _info_sufficient = False
        if self.path == '':
            _info_sufficient = False
        if self.crc == '':
            _info_sufficient = False
        return _info_sufficient


class DevicesTasks:
    def __init__(self, serial, timeout):
        self.serial = serial
        self.attempt = 1
        self.state = 0
        self.timeout = timeout
        self.ip_port = ''
        self.channel = ''
        self.channel_list = []
        self.rx_accept = ''
        self.log = []
        self.ftp_user = ''
        self.path2_choosen = False

    fw_states_names = {11: 'FwStart Command Sending',
                       2: 'FwStart Response Waiting',
                       3: 'Reconnection Waiting',
                       4: 'FwCheck Command Sending',
                       5: 'FwCheck Response Waiting',
                       6: 'Version Command Sending',
                       7: 'Version Response Waiting',
                       1: 'Version Command Sending',
                       12: 'Version Response Waiting',
                       20: 'Firmware Update OK',
                       21: 'Firmware Update Fail (Reconnection Timeout Expired)',
                       22: 'Firmware Update Fail (Attempts Expired)',
                       23: 'Firmware Update Canceled (There is no suitable one)',
                       24: 'Firmware Update Canceled (Versions Fully Match)'
                       }

    states_names = {1: 'Command Sending',
                    2: 'Response Waiting',
                    20: 'Response Received OK',
                    22: 'Response Received Fail (Attempts Expired)',
                    }

    in_progress = False
    channel = []
    command = ''
    timeout = 30
    timeout_reboot = 700
    attempts = 1
    parallels = 1
    tasks = []
    tasks_in_progress = 0
    global tx_queue_buffer

    @classmethod
    def processing(cls):
        global client_buffer

        def attempts_check(task):
            if task.attempt != cls.attempts:
                task.attempt += 1
                return True
            else:
                task.state = 22
                return False

        def timeout_attempts_check(task):
            if task.timeout == 0:
                return attempts_check(task)
            else:
                task.timeout -= 1
                return False

        def state_change(task, state, timeout=cls.timeout):
            task.state = state
            task.timeout = timeout

        def client_buffer_append(ip_port):
            _ip_port_find = False
            for _item in client_buffer:
                if _item[0] == ip_port:
                    _ip_port_find = True
                    break
            if not _ip_port_find:
                client_buffer.append([ip_port])

        def ip_port_get(task):
            _ip_port_state = 0  # 0 - not valid, 1 - valid, 2 - valid and changed from previous
            _ip_port = task.ip_port
            task.ip_port = ''
            for _device in DeviceTree:
                if _device.serial == task.serial:
                    if type(cls.channel) is str:
                        task.channel = cls.channel
                        for i in reversed(range(len(_device.channel))):
                            if _device.channel[i] == task.channel:
                                task.ip_port = _device.ip[i]
                                break
                    else:  # Проверить логику---------------------------
                        if task.state == 1:
                            # print('\n')
                            for _try in range(3):
                                if task.channel == '':
                                    task.channel = cls.channel[0]
                                    # print(f'{task.serial}  {task.channel}')
                                else:
                                    for i in range(len(cls.channel)):
                                        if task.channel == cls.channel[i]:
                                            _index = i
                                            if _index == len(cls.channel) - 1:
                                                task.channel = cls.channel[0]
                                            else:
                                                task.channel = cls.channel[_index + 1]
                                            # print(f'{task.serial}  {task.channel}')
                                            break

                                for i in reversed(range(len(_device.channel))):
                                    if _device.channel[i] == task.channel:
                                        task.ip_port = _device.ip[i]
                                        break
                                if task.ip_port != '':
                                    break
                        else:
                            for i in reversed(range(len(_device.channel))):
                                if _device.channel[i] == task.channel:
                                    task.ip_port = _device.ip[i]
                                    break
                    break

            if task.ip_port != '':  # Проверить логику---------------------------
                _ip_port_state += 1
                if _ip_port != task.ip_port:
                    _ip_port_state += 1
                    client_buffer_append(task.ip_port)

            if _ip_port != '' and _ip_port != task.ip_port:
                for i in range(len(client_buffer)):
                    if client_buffer[i][0] == _ip_port:
                        client_buffer.pop(i)
                        break

            return _ip_port_state

        # ------------------------------------------------------------------------------------------------------
        if not cls.in_progress:
            return

        for _task in cls.tasks:

            if cls.command.find('firmwareupdate#start#') >= 0:

                # ---------------------------------Version Command Send (matching station and firmware version)
                if _task.state == 1:
                    if ip_port_get(_task) > 0:
                        tx_queue_buffer.append([_task.ip_port, 'version\n'])
                        state_change(_task, 12)
                    if timeout_attempts_check(_task):
                        state_change(_task, 1)
                # ---------------------------------Version Command Response Wait (matching station and firmware version)
                elif _task.state == 12:
                    _task.rx_accept = f'version#OK#SoftwareVersion={fw_update_info.path[:2]}'
                    _log_get = cls.log_get(_task, False)
                    _path = fw_update_info.path
                    if _log_get == 1 and fw_update_info.path2 != '':
                        _task.path2_choosen = True
                        _task.rx_accept = f'version#OK#SoftwareVersion={fw_update_info.path2[:2]}'
                        _log_get = cls.log_get(_task, False)
                        _path = fw_update_info.path2

                    _task.rx_accept = f'version#OK#SoftwareVersion={_path[:_path.find(".Core")]}'
                    _log_get2 = cls.log_get(_task)

                    if _log_get == 2:
                        if _log_get2 == 2:
                            _task.state = 24
                        else:
                            _task.state = 11
                    elif _log_get == 1:
                        _task.state = 23
                    elif timeout_attempts_check(_task):
                        state_change(_task, 1)

                # ---------------------------------FW Start Command Send
                elif _task.state == 11:
                    if _task.ftp_user == '':
                        _task.ftp_user = FwUpdateInfo.ftp_users_rent()
                    if _task.ftp_user != '':
                        if ip_port_get(_task) > 0:

                            if _task.path2_choosen:
                                _command = f'{fw_update_info.path2}#crc={fw_update_info.crc2}'
                            else:
                                _command = f'{fw_update_info.path}#crc={fw_update_info.crc}'
                            _command = f'firmwareupdate#start#path={_command}#server={fw_update_info.server}' \
                                       f'#port={fw_update_info.port}#login={_task.ftp_user}' \
                                       f'#password={FwUpdateInfo.ftp_users_pass}#channel={_task.channel}\n'
                            tx_queue_buffer.append([_task.ip_port, _command])

                            _task.rx_accept = 'firmwareupdate#OK#Starting'
                            state_change(_task, 2)
                        if timeout_attempts_check(_task):
                            state_change(_task, 11)
                # ---------------------------------FW Start Command Response Wait
                elif _task.state == 2:
                    if cls.log_get(_task) == 2:
                        state_change(_task, 3, cls.timeout_reboot)
                    if timeout_attempts_check(_task):
                        state_change(_task, 11)
                # ---------------------------------Reconnection Wait
                elif _task.state == 3:
                    cls.log_get(_task)
                    if ip_port_get(_task) == 2:
                        state_change(_task, 4)
                    if timeout_attempts_check(_task):
                        _task.state = 21
                # ---------------------------------FW Check Command Send
                elif _task.state == 4:
                    if ip_port_get(_task) > 0:
                        tx_queue_buffer.append([_task.ip_port, 'firmwareupdate\n'])
                        _task.rx_accept = 'firmwareupdate#OK#StartupError=0 LoaderError=0 DownloaderError=0'
                        state_change(_task, 5)
                    if timeout_attempts_check(_task):
                        state_change(_task, 4)
                # ---------------------------------FW Check Command Response Wait
                elif _task.state == 5:
                    _log_get = cls.log_get(_task)
                    if _log_get == 2:
                        state_change(_task, 6)
                    elif _log_get == 1:
                        if attempts_check(_task):
                            state_change(_task, 11)
                    elif timeout_attempts_check(_task):
                        state_change(_task, 4)
                # ---------------------------------Version Command Send
                elif _task.state == 6:
                    if ip_port_get(_task) > 0:
                        tx_queue_buffer.append([_task.ip_port, 'version\n'])
                        if _task.path2_choosen:
                            _path = fw_update_info.path2
                        else:
                            _path = fw_update_info.path
                        _task.rx_accept = f'version#OK#SoftwareVersion={_path.replace(".Core.bin", "")}'
                        state_change(_task, 7)
                    if timeout_attempts_check(_task):
                        state_change(_task, 6)
                # ---------------------------------Version Command Response Wait
                elif _task.state == 7:
                    _log_get = cls.log_get(_task)
                    if _log_get == 2:
                        _task.state = 20
                    elif _log_get == 1:
                        if attempts_check(_task):
                            state_change(_task, 11)
                    elif timeout_attempts_check(_task):
                        state_change(_task, 6)

                # ---------------------------------Finish State
                if _task.state >= 20:
                    if _task.ftp_user != '':
                        FwUpdateInfo.ftp_users_release(_task.ftp_user)
                        _task.ftp_user = ''
                    for i in range(len(client_buffer)):
                        if client_buffer[i][0] == _task.ip_port:
                            client_buffer.pop(i)
                            break
            else:
                if _task.state == 1:
                    if ip_port_get(_task) > 0:
                        tx_queue_buffer.append([_task.ip_port, f'{cls.command}\n'])
                        if cls.command.find('#') > 0:
                            _task.rx_accept = cls.command[:cls.command.find('#')]
                        else:
                            _task.rx_accept = cls.command
                        _task.rx_accept += '#'
                        state_change(_task, 2)
                    if timeout_attempts_check(_task):
                        state_change(_task, 1)

                elif _task.state == 2:
                    if cls.log_get(_task) == 2:
                        _task.state = 20
                    if timeout_attempts_check(_task):
                        state_change(_task, 1)

                if _task.state >= 20:
                    for i in range(len(client_buffer)):
                        if client_buffer[i][0] == _task.ip_port:
                            client_buffer.pop(i)
                            break

        cls.log_update()
        cls.manager()

    @classmethod
    def group_start(cls):
        cls.tasks_in_progress = 0
        cls.in_progress = True
        cls.tasks = []
        cls.attempts = int(en_attempts.get())
        cls.parallels = int(en_parallel.get())
        FwUpdateInfo.ftp_users_init()
        cls.command = f'{en_command_group.get()}'
        if combo_group_channel.get() == 'auto':
            cls.channel = ['ethernet', 'wifi', 'gsm']
        else:
            cls.channel = combo_group_channel.get()
        _list = TreeDevices.get_children()
        for _item in _list:
            _serial = TreeDevices.item(_item, option="text")
            cls.tasks.append(DevicesTasks(_serial, cls.timeout))

    @classmethod
    def manager(cls):
        _parallel = 0
        for _task in cls.tasks:
            if 0 < _task.state < 20:
                _parallel += 1

        for _task in cls.tasks:
            if _parallel == cls.parallels:
                break
            if _task.state == 0:
                _parallel += 1
                _task.state = 1

        if _parallel == 0:
            window.after(10, group_toggle)

    @classmethod
    def log_get(cls, task, pop_item=True):
        _rx_accept_find = 0
        global client_buffer
        for _item in client_buffer:
            if _item[0] == task.ip_port:
                for i in range(1, len(_item), 1):
                    if _item[i].find('  TX:  ') >= 0:
                        continue
                    if _item[i].find(task.rx_accept[:task.rx_accept.find('#')]) >= 0:
                        if pop_item:
                            task.log.append(_item[i])
                        _rx_accept_find = 1
                        if _item[i].find(task.rx_accept) >= 0:
                            _rx_accept_find = 2
                if pop_item:
                    for i in reversed(range(1, len(_item), 1)):
                        _item.pop(i)
                break
        return _rx_accept_find

    @classmethod
    def log_update(cls):
        log_pos = log_fw.yview()
        if f'{log_pos[0]}' != '0.0':
            return
        log_fw.configure(state='normal')
        log_fw.delete(1.0, END)
        _count = 0
        for _task in cls.tasks:
            if _task.state != 0:
                _count += 1
                if cls.command.find('firmwareupdate#') >= 0:
                    _state_name = cls.fw_states_names[_task.state]
                else:
                    _state_name = cls.states_names[_task.state]
                log_fw.insert(END, f'{_task.serial:<13} Attempt = {_task.attempt:<2} Timeout = {_task.timeout:<4} '
                                   f'Channel = {_task.channel.replace("ernet", ""):<5} {_task.ftp_user:<5} '
                                   f'State = {_state_name}\n')
        cls.tasks_in_progress = _count
        log_fw.insert(END, f'\n')
        for _task in cls.tasks:
            if _task.state != 0:
                log_fw.insert(END, f'\n{_task.serial}')
                log_fw.insert(END, f'--------------------------------------------------------------------------------'
                                   f'------------------------------------\n')
                for _log in _task.log:
                    log_fw.insert(END, f'{_log[_log.find("#cmd#") + 4:].replace("<LF>##<LF>", "")}\n\n')
        # log_fw.yview(1.0)
        # log_fw.yview_moveto(f'{log_pos[0]:.2f}')
        log_fw.configure(state='disabled')


# --------------------------------------------------------------------------------------------------------------------
class ModbusMaster:
    crc_hi_tab = [0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                  0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                  0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
                  0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                  0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81,
                  0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0,
                  0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
                  0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                  0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                  0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                  0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
                  0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                  0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                  0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                  0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
                  0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                  0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                  0x40]
    crc_lo_tab = [0x00, 0xC0, 0xC1, 0x01, 0xC3, 0x03, 0x02, 0xC2, 0xC6, 0x06, 0x07, 0xC7, 0x05, 0xC5, 0xC4,
                  0x04, 0xCC, 0x0C, 0x0D, 0xCD, 0x0F, 0xCF, 0xCE, 0x0E, 0x0A, 0xCA, 0xCB, 0x0B, 0xC9, 0x09,
                  0x08, 0xC8, 0xD8, 0x18, 0x19, 0xD9, 0x1B, 0xDB, 0xDA, 0x1A, 0x1E, 0xDE, 0xDF, 0x1F, 0xDD,
                  0x1D, 0x1C, 0xDC, 0x14, 0xD4, 0xD5, 0x15, 0xD7, 0x17, 0x16, 0xD6, 0xD2, 0x12, 0x13, 0xD3,
                  0x11, 0xD1, 0xD0, 0x10, 0xF0, 0x30, 0x31, 0xF1, 0x33, 0xF3, 0xF2, 0x32, 0x36, 0xF6, 0xF7,
                  0x37, 0xF5, 0x35, 0x34, 0xF4, 0x3C, 0xFC, 0xFD, 0x3D, 0xFF, 0x3F, 0x3E, 0xFE, 0xFA, 0x3A,
                  0x3B, 0xFB, 0x39, 0xF9, 0xF8, 0x38, 0x28, 0xE8, 0xE9, 0x29, 0xEB, 0x2B, 0x2A, 0xEA, 0xEE,
                  0x2E, 0x2F, 0xEF, 0x2D, 0xED, 0xEC, 0x2C, 0xE4, 0x24, 0x25, 0xE5, 0x27, 0xE7, 0xE6, 0x26,
                  0x22, 0xE2, 0xE3, 0x23, 0xE1, 0x21, 0x20, 0xE0, 0xA0, 0x60, 0x61, 0xA1, 0x63, 0xA3, 0xA2,
                  0x62, 0x66, 0xA6, 0xA7, 0x67, 0xA5, 0x65, 0x64, 0xA4, 0x6C, 0xAC, 0xAD, 0x6D, 0xAF, 0x6F,
                  0x6E, 0xAE, 0xAA, 0x6A, 0x6B, 0xAB, 0x69, 0xA9, 0xA8, 0x68, 0x78, 0xB8, 0xB9, 0x79, 0xBB,
                  0x7B, 0x7A, 0xBA, 0xBE, 0x7E, 0x7F, 0xBF, 0x7D, 0xBD, 0xBC, 0x7C, 0xB4, 0x74, 0x75, 0xB5,
                  0x77, 0xB7, 0xB6, 0x76, 0x72, 0xB2, 0xB3, 0x73, 0xB1, 0x71, 0x70, 0xB0, 0x50, 0x90, 0x91,
                  0x51, 0x93, 0x53, 0x52, 0x92, 0x96, 0x56, 0x57, 0x97, 0x55, 0x95, 0x94, 0x54, 0x9C, 0x5C,
                  0x5D, 0x9D, 0x5F, 0x9F, 0x9E, 0x5E, 0x5A, 0x9A, 0x9B, 0x5B, 0x99, 0x59, 0x58, 0x98, 0x88,
                  0x48, 0x49, 0x89, 0x4B, 0x8B, 0x8A, 0x4A, 0x4E, 0x8E, 0x8F, 0x4F, 0x8D, 0x4D, 0x4C, 0x8C,
                  0x44, 0x84, 0x85, 0x45, 0x87, 0x47, 0x46, 0x86, 0x82, 0x42, 0x43, 0x83, 0x41, 0x81, 0x80,
                  0x40]
    test_data = '01 03 12 00 80 00 00 85 19 00 80 00 0D 00 01 00 7F 00 00 00 00 C0 2E'
    slave_address = 0
    slave_function = 0
    slave_register = 0

    @classmethod
    def crc_gen(cls, data):
        _crc_hi = 0xFF
        _crc_lo = 0xFF
        for _byte in data.split():
            _byte = int(_byte, 16)
            _index = _crc_hi ^ _byte
            _crc_hi = _crc_lo ^ cls.crc_hi_tab[_index]
            _crc_lo = cls.crc_lo_tab[_index]
        return f'{hex(_crc_hi).replace("0x", "").zfill(2)} {hex(_crc_lo).replace("0x", "").zfill(2)}'

    @classmethod
    def read_holding_reg(cls, address, register, quantity):
        cls.slave_address = address
        cls.slave_function = 3
        cls.slave_register = int(register, 16)
        register = f'{register:0>4}'
        register = f'{register[:2]} {register[2:]}'
        quantity = f'{quantity:0>4X}'
        quantity = f'{quantity[:2]} {quantity[2:]}'
        _data = f'{address:0>2X} 03 {register} {quantity}'
        _data = bytes.fromhex(f'{_data} {cls.crc_gen(_data)}')
        return _data

    @classmethod
    def parse_request(cls, data):
        pass  # можно засунуть в команду parse

    @classmethod
    def parse(cls, data):  # нужна проверка CRC и проверска на правильную длину
        # cls.read_holding_reg(1, '101', 9)
        _register_first = -1
        _register_last = -1
        _data = data.replace(' ', '')
        try:
            _quantity = int(_data[4:6], 16) // 2
        except ValueError:
            _quantity = 0
        _register = cls.slave_register
        _index = 6
        for reg in Solar.regs.keys():
            while True:
                if _quantity == 0 or reg <= _register:  # нужна ли проверка на < ???
                    break
                else:
                    _index += 4
                    _register += 1
                    _quantity -= 1

            if reg == _register:
                _type = Solar.regs[reg]['Type'].split('#')
                _data_type = _type[0]
                _data_lang = int(_type[1]) // 16

                _quantity -= _data_lang
                if _quantity < 0:
                    break

                if _register_first < 0:
                    _register_first = reg
                _register_last = reg

                _value = _data[_index:(_index + _data_lang * 4)]

                if _data_type == 'uint':
                    _value = int(_value, 16)
                    if Solar.regs[reg]['Scale'] != 1:
                        _value = float('{:.2f}'.format(_value * Solar.regs[reg]['Scale']))
                    Solar.regs[reg]['Value'] = _value
                elif _data_type == 'multipart':
                    _index1 = 0
                    for part in Solar.regs[reg].keys():
                        if type(part) is not str:
                            _type1 = Solar.regs[reg][part]['Type'].split('#')
                            _data_type1 = _type1[0]
                            _data_lang1 = int(_type1[1]) // 8
                            _value1 = _value[_index1:(_index1 + _data_lang1 * 2)]
                            _index1 += _data_lang1 * 2
                            if _data_type1 == '(sbit)uint':
                                _value1 = int(_value1, 16)
                                if _value1 - 128 >= 0:
                                    _value1 = (_value1 - 128) * (-1)
                                if Solar.regs[reg][part]['Scale'] != 1:
                                    _value1 = float('{:.2f}'.format(_value1 * int(Solar.regs[reg][part]['Scale'])))
                            Solar.regs[reg][part]['Value'] = _value1

                _register += _data_lang
                _index += _data_lang * 4

        return [_register_first, _register_last]


# --------------------------------------------------------------------------------------------------------------------
class Solar:
    regs = {256: {'Name': 'BatteryCapacity', 'Type': 'uint#16', 'Scale': 1, 'Unit': '%', 'Value': ''},
            257: {'Name': 'BatteryVoltage', 'Type': 'uint#16', 'Scale': 0.1, 'Unit': 'V', 'Value': ''},
            258: {'Name': 'BatteryChargingCurrent', 'Type': 'uint#16', 'Scale': 0.01, 'Unit': 'A', 'Value': ''},
            259: {'Type': 'multipart#16',
                  1: {'Name': 'ControllerTemperature', 'Type': '(sbit)uint#8', 'Scale': 1, 'Unit': 'C', 'Value': ''},
                  2: {'Name': 'BatteryTemperature', 'Type': '(sbit)uint#8', 'Scale': 1, 'Unit': 'C', 'Value': ''},
                  },
            260: {'Name': 'LoadVoltage', 'Type': 'uint#16', 'Scale': 0.1, 'Unit': 'V', 'Value': ''},
            261: {'Name': 'LoadCurrent', 'Type': 'uint#16', 'Scale': 0.01, 'Unit': 'A', 'Value': ''},
            262: {'Name': 'LoadPower', 'Type': 'uint#16', 'Scale': 1, 'Unit': 'W', 'Value': ''},
            263: {'Name': 'SolarVoltage', 'Type': 'uint#16', 'Scale': 0.1, 'Unit': 'V', 'Value': ''},
            264: {'Name': 'SolarCurrent', 'Type': 'uint#16', 'Scale': 0.01, 'Unit': 'A', 'Value': ''},
            265: {'Name': 'SolarPower', 'Type': 'uint#16', 'Scale': 1, 'Unit': 'W', 'Value': ''},

            277: {'Name': 'OperatingDays', 'Type': 'uint#16', 'Scale': 1, 'Unit': '', 'Value': ''},
            278: {'Name': 'BatteryDischargesCount', 'Type': 'uint#16', 'Scale': 1, 'Unit': '', 'Value': ''},
            279: {'Name': 'BatteryFullChargesCount', 'Type': 'uint#16', 'Scale': 1, 'Unit': '', 'Value': ''},

            284: {'Name': 'PowerGeneration', 'Type': 'uint#32', 'Scale': 1, 'Unit': 'W', 'Value': ''},
            286: {'Name': 'PowerConsumption', 'Type': 'uint#32', 'Scale': 1, 'Unit': 'W', 'Value': ''},
            }

    days_current = 0
    days_quantity = 0
    serial = 'MPPT2420'
    cli_strings = [serial, 'AT+CIP', 'channelinfo']

    @classmethod
    def command_prepare(cls, data):
        if data == 'history':
            data = ModbusMaster.read_holding_reg(1, '115', 10)
        elif data.find('history#daysquantity=') >= 0:
            try:
                _quantity = int(data[len('history#daysquantity='):], 16)
            except ValueError:
                pass
            else:
                cls.days_quantity = _quantity
                cls.days_current = 0

        elif data == 'history#drop':
            data = bytes.fromhex('01 79 00 00 00 01 5D C0')
        elif data == 'measure':
            data = ModbusMaster.read_holding_reg(1, '100', 10)
        elif data == 'settings':
            pass
        elif data.find('AT+CIP') == 0:
            data += '\r\n'
        elif data == 'channelinfo' == 0:
            data += '\n'

        return data


# --------------------------------------------------------------------------------------------------------------------
class Devices:
    def __init__(self, obj, ip_port, time=''):
        self.obj = obj
        self.ip_port = ip_port
        self.time = time
        self.serial = 'Unknown'
        self.channel = ''
        self.timeout = 60
        self.data1 = []
        self.data2 = []
        self.buffer_data1_active = True
        self.buffer_to_file = False
        self.delete_state = False

    # def __del__(self):
    # print("deleted")

    data_buffer = 20  # In Lines

    def data_append(self, data, direction):
        _data_is_hex = ''
        if type(data) == bytes:
            _data_is_hex = '(hex)'
            _data = data
            try:
                data = data.decode('windows-1251')
            except ValueError:
                pass
            else:
                if self.serial.find(Solar.serial) == 0:
                    for _string in Solar.cli_strings:
                        if data.find(_string) >= 0:
                            _data_is_hex = ''
                            break
                else:
                    _data_is_hex = ''
            if _data_is_hex != '':
                data = _data.hex(' ', 1).upper()

        if _data_is_hex == '':
            if direction == 'RX':
                self.ser_ch_define(data)
            data = data.replace('\n', '<LF>')
            data = data.replace('\r', '<CR>')

        _time = datetime.datetime.now()
        _time = _time.strftime('%Y.%m.%d %H:%M:%S.') + str(int(_time.microsecond / 1000)).zfill(3)
        data = f'{_time}  {direction}{_data_is_hex}:  {data}'

        global client_buffer
        for _item in client_buffer:
            if _item[0] == self.ip_port:
                _item.append(data)

        global tree_chosen
        if tree_chosen != '':
            if TreeDevices.item(tree_chosen, option="values")[0] == self.ip_port:
                datalog_insert(data)
                if TreeDevices.item(TreeDevices.parent(tree_chosen), option="text").find(Solar.serial) == 0:
                    if direction == 'RX' and _data_is_hex == '(hex)':
                        _registers = ModbusMaster.parse(data[35:])
                        if _registers[0] > 0:
                            data = ''
                            for reg in Solar.regs.keys():
                                if reg > _registers[1]:
                                    break
                                if reg >= _registers[0]:
                                    if Solar.regs[reg]['Type'].find('multipart') >= 0:
                                        for part in Solar.regs[reg].keys():
                                            if type(part) is not str:
                                                data += f"{'':^35s}{Solar.regs[reg][part]['Name']} = " \
                                                        f"{Solar.regs[reg][part]['Value']} " \
                                                        f"{Solar.regs[reg][part]['Unit']}\n"
                                    else:
                                        data += f'{"":^35s}{Solar.regs[reg]["Name"]} = ' \
                                                f'{Solar.regs[reg]["Value"]} ' \
                                                f'{Solar.regs[reg]["Unit"]}\n'
                            datalog_insert(data)

        if self.buffer_data1_active is True:
            if len(self.data1) < Devices.data_buffer:
                self.data1.append(data)
            else:
                self.buffer_data1_active = False
                self.buffer_to_file = True
                self.data2 = []
                self.data2.append(data)
        else:
            if len(self.data2) < Devices.data_buffer:
                self.data2.append(data)
            else:
                self.buffer_data1_active = True
                self.buffer_to_file = True
                self.data1 = []
                self.data1.append(data)

    def ser_ch_define(self, _data):
        if self.serial == 'Unknown':
            _serial = ''
            _channel = ''
            if _data[0] == '#':
                _index = _data.find('\n#cmd#')
                if 2 <= _index <= 13:
                    _serial = _data[1:_index]
                    _index = _data.find('InterfaceType=')
                    if _index >= 0:
                        _index = _data.find('\"', _index) + 1
                        _index2 = _data.find('\"', _index)
                        if _index2 >= 0:
                            _channel = _data[_index:_index2].lower()
            # сделано так, лишь потому что если первый RX был не ченелинфо в тривью ченел не обновится и будет пустым
            if _serial != '' and _channel != '':
                self.serial = _serial
                self.channel = _channel

    def allowed_check(self):
        _allowed = True
        if self.serial == 'Unknown':
            if self.timeout > 0:
                self.timeout -= 1
            elif self.timeout == 0:
                _allowed = False
        return _allowed


class DevicesTree:
    def __init__(self, serial):
        self.serial = serial
        self.ser_obj = ''
        self.ip = []
        self.channel = []
        self.con_obj = []


# ------------------------------------------------------------------------------------------------------
def devices_from_file(event):
    global device_lists
    try:
        file = open('DeviceList.txt', 'r', encoding='windows-1251').readlines()
    except OSError:
        pass
    else:
        device_lists = {}
        list_current = ''
        for line in file:
            _line = line.replace('\n', '')
            if _line[0:1] == '#':
                list_current = _line[1:]
                device_lists.update({list_current: []})
            else:
                if _line != '':
                    device_lists[list_current].append(_line)
    combo_values = []
    for key in device_lists.keys():
        combo_values.append(key)
    combo_filter['values'] = combo_values


# ------------------------------------------------------------------------------------------------------
def write_to_file_init():
    try:
        os.mkdir(f'{os.getcwd()}\\Logs')
    except FileExistsError:
        pass
    except OSError:
        pass


def write_to_file(i):
    Device[i].buffer_to_file = False
    if Device[i].buffer_data1_active is True:
        _data = Device[i].data2
    else:
        _data = Device[i].data1

    if len(_data) != 0:
        _file_dir = f'{os.getcwd()}\\Logs\\{Device[i].serial}\\'
        _file_name = f'{Device[i].time}_{Device[i].serial}_{Device[i].channel}' \
                     f'_{Device[i].ip_port.replace(":", "_")}.log'
        try:
            os.mkdir(_file_dir)
        except FileExistsError:
            pass
        except OSError:
            pass
        try:
            file = open(f'{_file_dir}{_file_name}', 'a')
        except OSError:
            pass
        else:
            for _row in _data:
                file.write(f'{_row}\n')
            file.close()  # нужна проверка на наличие свободного места, либо try или with


# ----------CleanUp DataBuffers (Buffers to Files), CleanUp Devices-------------------------------
def clients_cleanup():
    global count_3d2
    global count_mppt
    global tx_queue_buffer
    while True:
        for i in reversed(range(len(Device))):
            if count_3d2 == 10:
                if Device[i].serial == 'CA01PM0003D2':
                    tx_queue_buffer.append([Device[i].ip_port, 'gsm\n'])

            if count_mppt == 50:
                if Device[i].serial.find(Solar.serial) == 0:
                    tx_queue_buffer.append([Device[i].ip_port, 'channelinfo\n'])

            if not Device[i].allowed_check():
                Device[i].delete_state = True

            if Device[i].buffer_to_file is True:
                write_to_file(i)

            if Device[i].delete_state is True:
                Device[i].buffer_data1_active = not Device[i].buffer_data1_active
                write_to_file(i)

                try:
                    Device[i].obj.close()
                except OSError:
                    pass
                if tree_chosen != '':
                    if TreeDevices.item(tree_chosen, option="values")[0] == Device[i].ip_port:
                        datalog_insert('<CLOSED>')
                Device.pop(i)

        if count_mppt == 50:
            count_mppt = 0
        else:
            count_mppt += 1

        if count_3d2 == 10:
            count_3d2 = 0
        else:
            count_3d2 += 1

        window.after(10, DevicesTasks.processing)
        time.sleep(1)


# ------------------------------------------------------------------------------------------------------
def clients_check():
    global timing_en
    while True:
        if timing_en:
            print('\n')
            time1 = datetime.datetime.now().microsecond

        for i in range(100):
            try:
                _obj, _daddr = tcpServ.accept()
            except OSError:
                pass
            else:
                try:
                    _obj.send(b'channelinfo\x0d\x0a')
                except OSError:
                    pass
                _ip_port = f'{_daddr[0]}:{_daddr[1]}'
                _time = datetime.datetime.now()
                Device.append(Devices(obj=_obj, ip_port=_ip_port, time=_time.strftime('%y%m%d_%H%M%S')))
                # print(f'{_daddr[1]}  ::  {len(Device)}')

        if timing_en:
            time2 = datetime.datetime.now().microsecond
            _time = time2 - time1
            if _time < 0:
                _time += 1000000
            print(f'{"Accept":15}{_time}')

        for i in reversed(range(len(Device))):
            try:
                _data = Device[i].obj.recv(4096)
            except BlockingIOError:
                pass
            except OSError:
                Device[i].delete_state = True
            else:
                if _data != b'':
                    Device[i].data_append(_data, 'RX')
                else:
                    Device[i].delete_state = True

            #
            # _data = Device[i].obj.recv(4096)
            # if _data != b'':
            #     Device[i].data_append(_data, 'RX')
            # else:
            #     Device[i].delete_state = True

            global tx_queue_buffer
            for j in range(len(tx_queue_buffer)):
                if Device[i].ip_port == tx_queue_buffer[j][0]:
                    try:
                        if type(tx_queue_buffer[j][1]) is bytes:
                            Device[i].obj.send(tx_queue_buffer[j][1])
                        else:
                            Device[i].obj.send(bytes(tx_queue_buffer[j][1].encode('windows-1251')))
                    except OSError:
                        pass
                    else:
                        Device[i].data_append(tx_queue_buffer[j][1], 'TX')
                    finally:
                        tx_queue_buffer.pop(j)
                    break

        if timing_en:
            time3 = datetime.datetime.now().microsecond
            _time = time3 - time2
            if _time < 0:
                _time += 1000000
            print(f'{"Rx/Tx":15}{_time}')

        window.after(5, tree_devices_update)
        time.sleep(0.1)

    # global tree_devices_update_count
    # print(f'treeUpdate: {tree_devices_update_count}')
    # if tree_devices_update_count == 2:
    #     window.after(1, tree_devices_update)
    #     tree_devices_update_count = 0
    # else:
    #     tree_devices_update_count += 1


# --------------------------------------Creating Connections List -------------------------------------
def tree_devices_update():
    global timing_en
    if timing_en:
        time1 = datetime.datetime.now().microsecond

    serial_from_file = ''
    for key in device_lists.keys():
        if combo_filter.get() == key:
            serial_from_file = key
            break

    _serial_list = []
    if serial_from_file != '':
        for _device in device_lists[key]:
            _serial_list.append(_device)
    else:
        for _device in Device:
            _serial_list.append(_device.serial)
        _serial_list = sorted(list(set(_serial_list)))

    _ip_list = []
    _channel_list = []
    for _serial in _serial_list:
        _ip = []
        _channel = []
        for _device in Device:
            if _device.serial == _serial:
                _ip.append(_device.ip_port)
                _channel.append(_device.channel)
        _ip_list.append(_ip)
        _channel_list.append(_channel)

    if serial_from_file != '':
        for i in reversed(range(len(_serial_list))):
            if len(_ip_list[i]) == 0:
                _serial_list.append(_serial_list.pop(i))
                _ip_list.append(_ip_list.pop(i))
                _channel_list.append(_channel_list.pop(i))

    if timing_en:
        time2 = datetime.datetime.now().microsecond
        _time = time2 - time1
        if _time < 0:
            _time += 1000000
        print(f'{"List":15}{_time}')

    # --------------------------------------Delete Dead Connections-------------------------------------
    global tree_chosen
    global filter_updated
    for i in reversed(range(len(DeviceTree))):
        _serial_find = False
        for i2 in range(len(_serial_list)):
            if DeviceTree[i].serial == _serial_list[i2]:
                for i3 in reversed(range(len(DeviceTree[i].ip))):
                    _ip_find = False
                    for i4 in range(len(_ip_list[i2])):
                        if DeviceTree[i].ip[i3] == _ip_list[i2][i4]:
                            _ip_find = True
                            break
                    if not _ip_find:
                        if DeviceTree[i].con_obj[i3] != tree_chosen:
                            TreeDevices.delete(DeviceTree[i].con_obj[i3])
                            DeviceTree[i].con_obj.pop(i3)
                            DeviceTree[i].channel.pop(i3)
                            DeviceTree[i].ip.pop(i3)
                            if len(_ip_list[i2]) == 0:
                                filter_updated = True
                _serial_find = True
                break
        # if not _serial_find:
        #     for _con_obj in DeviceTree[i].con_obj:
        #         if _con_obj == tree_focus:
        #             _serial_find = True
        if not _serial_find:
            for _con_obj in DeviceTree[i].con_obj:
                if _con_obj == tree_chosen:
                    tree_chosen = ''
                    break
            TreeDevices.delete(DeviceTree[i].ser_obj)
            DeviceTree.pop(i)

    # --------------------------------------Append New Connections-------------------------------------
    for i in range(len(_serial_list)):
        _serial_find = False
        for i2 in range(len(DeviceTree)):
            if _serial_list[i] == DeviceTree[i2].serial:
                for i3 in range(len(_ip_list[i])):
                    _ip_find = False
                    for i4 in range(len(DeviceTree[i2].ip)):
                        if _ip_list[i][i3] == DeviceTree[i2].ip[i4]:
                            _ip_find = True
                            break
                    if not _ip_find:
                        DeviceTree[i2].con_obj.append(TreeDevices.insert(DeviceTree[i2].ser_obj, END,
                                                                         text=_channel_list[i][i3],
                                                                         values=_ip_list[i][i3]))
                        DeviceTree[i2].channel.append(_channel_list[i][i3])
                        DeviceTree[i2].ip.append(_ip_list[i][i3])
                        filter_updated = True
                _serial_find = True
                break
        if not _serial_find:
            DeviceTree.append(DevicesTree(_serial_list[i]))
            _index = len(DeviceTree) - 1
            DeviceTree[_index].ser_obj = TreeDevices.insert('', END, text=DeviceTree[_index].serial)
            TreeDevices.detach(DeviceTree[_index].ser_obj)
            filter_updated = True
            for i2 in range(len(_ip_list[i])):
                DeviceTree[_index].con_obj.append(TreeDevices.insert(DeviceTree[_index].ser_obj, END,
                                                                     text=_channel_list[i][i2],
                                                                     values=_ip_list[i][i2]))
                DeviceTree[_index].channel.append(_channel_list[i][i2])
                DeviceTree[_index].ip.append(_ip_list[i][i2])

    if timing_en:
        time3 = datetime.datetime.now().microsecond
        _time = time3 - time2
        if _time < 0:
            _time += 1000000
        print(f'{"Tree":15}{_time}')

    # return

    # --------------------------------------Filtering/Sorting Connections-------------------------------------
    if filter_updated:
        filter_updated = False

        if timing_en:
            time1 = datetime.datetime.now().microsecond

        _serial_list_filtered = []
        if serial_from_file != '':
            _serial_list_filtered = _serial_list
        else:
            _filters = combo_filter.get().replace(' ', '').split(',')
            for _serial in _serial_list:
                _filter_pass = True
                for _filter in _filters:
                    if _filter == '':
                        continue
                    _filter_pass = False
                    if _serial.find(_filter) >= 0 or _serial.find(_filter.upper()) >= 0:
                        _filter_pass = True
                        break
                if _filter_pass:
                    _serial_list_filtered.append(_serial)

        global tree_expand_all
        global tree_collapse_all
        if tree_expand_new.get():
            tree_expand_all = True
        for _device in DeviceTree:
            if len(_serial_list_filtered) == 0:
                TreeDevices.detach(_device.ser_obj)
            for i in range(0, len(_serial_list_filtered), 1):
                if _device.serial == _serial_list_filtered[i]:
                    TreeDevices.reattach(_device.ser_obj, '', i)
                    if tree_expand_all is True:
                        TreeDevices.item(_device.ser_obj, open=True)
                    if tree_collapse_all is True:
                        TreeDevices.item(_device.ser_obj, open=False)
                    break
                if i == len(_serial_list_filtered) - 1:
                    TreeDevices.detach(_device.ser_obj)

        tree_expand_all = False
        tree_collapse_all = False

        if timing_en:
            time2 = datetime.datetime.now().microsecond
            _time = time2 - time1
            if _time < 0:
                _time += 1000000
            print(f'{"Filter":15}{_time}')


# ------------------------------------------------------------------------------------------------------
def tree_focus_set(event):
    global tree_chosen
    global tree_selected
    if tree_selected != TreeDevices.selection()[0]:
        try:
            TreeDevices.selection_remove(tree_selected)
        except TclError:
            pass
        tree_selected = TreeDevices.selection()[0]
    if tree_selected != tree_chosen:
        if len(TreeDevices.get_children(tree_selected)) == 0:
            if len(TreeDevices.item(tree_selected, option="values")) != 0:
                tree_chosen = tree_selected
                for _device in Device:
                    if _device.ip_port == TreeDevices.item(tree_chosen, option="values")[0]:
                        if _device.buffer_data1_active:
                            datalog_insert(_device.data2, append=False)
                            datalog_insert(_device.data1)
                        else:
                            datalog_insert(_device.data1, append=False)
                            datalog_insert(_device.data2)
                        break


def tree_clipboard_append(event):
    global tree_selected
    if len(TreeDevices.get_children(tree_selected)) != 0:
        _clipboard = TreeDevices.item(tree_selected, option="text")
    else:
        _clipboard = TreeDevices.item(tree_selected, option="values")[0]
    window.clipboard_clear()
    window.clipboard_append(_clipboard)


def combo_filter_event(event):
    global filter_updated
    filter_updated = True


def datalog_insert(data, append=True, decode=''):
    log_main.configure(state='normal')
    if append:
        _temp = log_main.index(END)
        _temp = int(_temp[0: _temp.find(',') - 1])
        if _temp > DATALOG_BUFFER:
            log_main.delete(1.0, str(_temp - DATALOG_BUFFER) + '.0')
    else:
        log_main.delete(1.0, END)
    if type(data) is list:
        for _data in data:
            log_main.insert(END, _data + '\n\n')
    else:
        log_main.insert(END, data + '\n\n')
    log_main.configure(state='disabled')
    log_main.yview(END)


def en_command_event():
    global tree_chosen
    global tx_queue_buffer
    if tree_chosen != '':
        _ip_port = TreeDevices.item(tree_chosen, option="values")[0]
        if TreeDevices.item(TreeDevices.parent(tree_chosen), option="text").find(Solar.serial) == 0:
            _data = Solar.command_prepare(en_command.get())
        else:
            _data = en_command.get().replace('#', ' #') + '\n'
        if len(tx_queue_buffer) < 10:
            tx_queue_buffer.append([_ip_port, _data])


def ftp_refresh():
    threading.Timer(0.01, ftp_connect).start()


def combo_fw_list_select_event(event):
    combo_fw_select('retr1')


def combo_fw_list2_select_event(event):
    combo_fw_select('retr2')


def combo_fw_select(mode='retr1'):
    threading.Timer(0.01, lambda: ftp_connect(mode=mode)).start()


def ftp_connect(mode=None):
    def crc(data):
        _crc = 0x0000
        for _block in data:
            _block = _block.hex(' ', 1)
            for _byte in _block.split():
                _byte = int(_byte, 16)
                _index = (_crc >> 8) ^ _byte
                _crc = ((_crc & 255) << 8) ^ crc_tab[_index]
        return int(_crc)

    if mode is not None:
        en_command_group.delete(0, END)
        en_command_group.insert(0, 'firmwareupdate#start')

    with FTP() as ftp:
        try:
            ftp.connect('cityair.uniscan.biz', 6658)
            ftp.login(user='FTP', passwd='password')
        except:
            return
        fw_update_info.server = 'cityair.uniscan.biz'
        fw_update_info.port = '6658'
        fw_update_info.login = 'FTP'
        fw_update_info.password = 'password'

        if mode is None:
            _fw_list = []
            ftp.retrlines('NLST', _fw_list.append)
            _fw_list_filtered = ['']
            for _fw_name in _fw_list:
                if _fw_name[_fw_name.rfind('.') + 1:] == 'bin':
                    _fw_list_filtered.append(_fw_name)
            combo_fw_list['values'] = _fw_list_filtered
            combo_fw_list2['values'] = _fw_list_filtered

        else:
            fw_update_info.path = combo_fw_list.get()
            if fw_update_info.path != '':
                _fw = []
                ftp.retrbinary(f'RETR {fw_update_info.path}', _fw.append)
                fw_update_info.crc = crc(_fw)
                fw_update_info.channel = combo_group_channel.get()

                combo_fw_list2.configure(state='readonly')
                fw_update_info.path2 = combo_fw_list2.get()
                if fw_update_info.path2 != '':
                    _fw2 = []
                    ftp.retrbinary(f'RETR {fw_update_info.path2}', _fw2.append)
                    fw_update_info.crc2 = crc(_fw2)
            else:
                combo_fw_list2.configure(state='disable')
            if mode == 'retr1':
                en_command_fw_update(fw_update_info.path, fw_update_info.crc)
            elif mode == 'retr2':
                en_command_fw_update(fw_update_info.path2, fw_update_info.crc2)
        ftp.quit()


def en_command_fw_update(path, crc):
    en_command_group.delete(0, END)
    en_command_group.insert(END, f'firmwareupdate#start'
                                 f'#path={path}#crc={crc}'
                                 f'#server={fw_update_info.server}#port={fw_update_info.port}'
                                 f'#login={fw_update_info.login}#password={fw_update_info.password}'
                                 f'#channel=')


def en_command_group_event():
    if btn_fw_start.cget('text') == 'Start':
        btn_fw_start.configure(text="Stop", fg='red')
        DevicesTasks.group_start()


def group_toggle():
    if btn_fw_start.cget('text') == 'Start':
        btn_fw_start.configure(text="Stop", fg='red')
        DevicesTasks.group_start()
    else:
        DevicesTasks.in_progress = False
        btn_fw_start.configure(text="Start", fg='black')


def tree_menu_open(event):
    TreeDevices_menu.post(TreeDevices.winfo_rootx() + event.x, TreeDevices.winfo_rooty() + event.y)


def tree_menu_expand_all():
    global tree_expand_all
    global filter_updated
    filter_updated = True
    tree_expand_all = True


def tree_menu_collapse_all():
    global tree_collapse_all
    global filter_updated
    filter_updated = True
    tree_collapse_all = True


def win_conf_event(event):
    global log_fw_win_fix
    en_command_group.win.withdraw()
    en_command.win.withdraw()
    log_fw_win_fix = False
    log_fw_win.withdraw()


def log_fw_win_update(x=0, y=2, moveto=0):
    global log_fw_win_y
    global log_fw_win_fix
    _count = DevicesTasks.tasks_in_progress
    _draw = False
    if moveto != 0:
        x = 5
        y = log_fw_win_y+4
        if moveto > 0:
            y = log_fw_win_y
    # print(f'{x}x{y}')
    if (0 <= x <= 100 and 2 <= y < log_fw.bbox(f'{_count + 1}.0')[1]) or log_fw_win_fix:
        i = 0
        for i in range(1, _count + 1, 1):
            # print(f'({y} {log_fw.bbox(str(i) + ".0")[1]}-{log_fw.bbox(str(i + 1) + ".0")[1]})')
            if log_fw.bbox(f'{i}.0')[1] <= y <= (log_fw.bbox(f'{i + 1}.0')[1] + 3):
                break
        # print(f'{i}:{_count}')
        if moveto < 0:
            pass
        elif moveto > 0:
            i -= 1
        elif moveto == 0:
            i -= 1

        y = log_fw.bbox(f'{i+1}.0')[1]
        # print(f'i+1 = {i}')
        if 0 <= i < _count:
            if moveto != 0:
                _draw = True
            log_fw_win_text.delete(1.0, END)
            log_fw_win_text.insert(END, DevicesTasks.tasks[i].serial[:6] + ' ' +
                                   DevicesTasks.tasks[i].serial[6:] + '\n\n')

            if DevicesTasks.command.find('firmwareupdate#start#') >= 0:
                for _log in DevicesTasks.tasks[i].log:
                    log_fw_win_text.insert(END, _log + '\n\n')
            else:
                for command in CmdExample.keys():
                    if len(DevicesTasks.tasks[i].log) == 0:
                        break
                    _log = DevicesTasks.tasks[i].log[0].replace('<LF>##<LF>', '\n')
                    if command == DevicesTasks.command.split('#')[0]:
                        try:
                            CmdExample[command]['Disable']
                        except KeyError:
                            parse_ca(_log, CmdExample)
                            _width = 0
                            for key in CmdExample[command].keys():
                                if len(key) > _width:
                                    _width = len(key)
                            for key in CmdExample[command].keys():
                                _type = CmdExample[command][key]['Type']
                                if _type == 'Entry' or _type == 'Combo' or _type == 'Label':
                                    log_fw_win_text.insert(END, key.ljust(_width + 2)
                                                           + CmdExample[command][key]["Value"] + '\n')
                        else:
                            if command == 'check':
                                _log = _log[_log.find('FailCount='):].replace('} {', '\n')
                                _log = _log.replace(' {', '\n').replace('}', '').replace('\"', '')
                                for _line in _log.split('\n'):
                                    _line = _line.split(', ')
                                    if len(_line) == 1:
                                        log_fw_win_text.insert(END, _line[0] + '\n')
                                    else:
                                        log_fw_win_text.insert(END, _line[0].ljust(16) + _line[1].ljust(7)
                                                               + _line[2].ljust(5) + _line[3] + '\n')
                            else:
                                log_fw_win_text.insert(END, _log[:_log.rfind('#') + 1] + '\n' +
                                                       _log[_log.rfind('#') + 1:])

                        break
            _draw = True
    if _draw:
        if log_fw_win_fix:
            log_fw_win_text.configure(bg='grey95', highlightbackground="grey", highlightthickness=5, padx=6, pady=6)
        else:
            log_fw_win_text.configure(bg='grey95', highlightbackground="grey", highlightthickness=1, padx=10, pady=10)
            log_fw_win.geometry(f'+{log_fw.winfo_rootx() + 105}+{log_fw.winfo_rooty() + y + 20}')
        log_fw_win.deiconify()
        log_fw_win_y = y
        # print(f'log_fw_win_y = {log_fw_win_y}')
    else:
        if moveto == 0:
            log_fw_win_fix = False
            log_fw_win.withdraw()


def win_button_event(event):
    global log_fw_win_fix
    global log_fw_win_y
    global win_motion_after
    if log_fw_win.state() == 'normal':
        win_motion()
        log_fw_win_fix = not log_fw_win_fix
        window.after_cancel(win_motion_after)
        win_motion_after = window.after(100, win_motion)
        # log_fw_win_update(y=log_fw_win_y)
        return 'break'


def win_motion():
    if Tabs.index('current') == 1:
        log_fw_win_update(window.winfo_pointerx() - log_fw.winfo_rootx(),
                          window.winfo_pointery() - log_fw.winfo_rooty())


def win_motion_event(event):
    global win_motion_after
    if not log_fw_win_fix:
        win_motion()
        window.after_cancel(win_motion_after)
        win_motion_after = window.after(300, win_motion)


def log_fw_wheel_event(event):
    if log_fw_win.state() == 'normal':
        log_fw_win_update(moveto=event.delta)
        return 'break'


if __name__ == '__main__':
    tcpServ = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    tcpServ.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
    tcpServ.ioctl(socket.SIO_KEEPALIVE_VALS, (1, 90 * 1000, 20 * 1000))  # on/off, timeout(ms), period(ms)
    # after timeout tries 3 attempts keepalive with period and then closes connection after time=period
    tcpServ.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    # tcpServ.setsockopt(socket.SOL_SOCKET, socket.SO_SNDBUF, 115200)
    # tcpServ.settimeout(0.0000001)
    tcpServ.setblocking(False)
    tcpServ.bind(('0.0.0.0', 6665))
    tcpServ.listen()
    client_obj = []
    client = []
    Device = []
    DeviceTree = []
    devices_tasks = []
    tree_devices_update_count = 0
    fw_update_info = FwUpdateInfo()

    window = Tk()
    window.title('DebugServer')
    window.geometry('1500x950+200+40')
    window.bind('<Configure>', win_conf_event)
    window.bind('<FocusOut>', win_conf_event)
    window.bind('<Escape>', win_conf_event)
    window.bind('<Motion>', win_motion_event)
    window.bind('<Button-2>', win_button_event)

    win_motion_after = window.after(100, win_motion)


    FrameDevices = Frame(window)
    FrameDevices.pack(side=LEFT, anchor=NW, padx=6, pady=3, fill=Y, expand=1)

    filter_updated = False
    combo_filter = Combobox(FrameDevices, width=41, height=40)
    combo_filter.pack(side=TOP, padx=0, pady=3, anchor=NW)
    combo_filter.bind('<Button-1>', devices_from_file)
    combo_filter.bind('<KeyRelease>', combo_filter_event)

    TreeDevices = Treeview(FrameDevices, show="tree", selectmode="browse")
    TreeDevices["columns"] = '#1'
    TreeDevices.column("#0", width=120, minwidth=20, stretch=NO)
    TreeDevices.column('#1', width=130, minwidth=20, stretch=NO)
    TreeDevices.pack(side=LEFT, fill=Y, padx=0, pady=3, expand=1)
    ScrollDevices = Scrollbar(FrameDevices, orient="vertical", command=TreeDevices.yview)
    ScrollDevices.pack(side=RIGHT, fill=Y)
    TreeDevices.configure(yscrollcommand=ScrollDevices.set)
    TreeDevices.bind('<Control-c>', tree_clipboard_append)
    TreeDevices.bind('<Control-C>', tree_clipboard_append)
    TreeDevices.bind('<<TreeviewSelect>>', tree_focus_set)
    TreeDevices.bind("<Button-3>", tree_menu_open)
    TreeDevices_menu = Menu(TreeDevices, tearoff=0)
    tree_expand_new = BooleanVar(value=False)
    TreeDevices_menu.add_checkbutton(label="Expand New", variable=tree_expand_new)
    TreeDevices_menu.add_command(label="Expand All", command=tree_menu_expand_all)
    TreeDevices_menu.add_command(label="Collapse All", command=tree_menu_collapse_all)

    Tabs = Notebook(window)
    tab_main = Frame(Tabs, width=300)
    tab_fw_update = Frame(Tabs, width=300)
    Tabs.add(tab_main, text=f'{"CLI": ^20s}')
    Tabs.add(tab_fw_update, text=f'{"Group CLI": ^20s}')
    Tabs.pack(side=RIGHT, expand=1, fill=BOTH, padx=3, pady=3)

    en_command = EntryAutoFill(tab_main, TOP, 'single')

    frame_log_main = Frame(tab_main)
    frame_log_main.pack(side=BOTTOM, anchor=NW, padx=6, pady=0)

    log_main = ScrolledText(master=frame_log_main, width=300, height=100, wrap=WORD)
    log_main.insert(END, '')
    log_main.configure(state='disabled')
    log_main.pack(side=LEFT)

    frame_fw_controls = Frame(tab_fw_update)
    frame_fw_controls.pack(side=TOP, anchor=NW, padx=0, pady=3, fill=X, expand=1)

    en_command_group = EntryAutoFill(frame_fw_controls, TOP, 'group')

    lbl = Label(frame_fw_controls, text=' Attempts =')
    lbl.pack(side=LEFT)

    en_attempts = Entry(frame_fw_controls, width=3, validate="key")
    # en_attempts ['validatecommand'] = (en_x_max.register(en_x_val), '%P', '%d')
    en_attempts.pack(side=LEFT)
    en_attempts.insert(END, '2')

    lbl = Label(frame_fw_controls, text=' Parallel =')
    lbl.pack(side=LEFT)

    en_parallel = Entry(frame_fw_controls, width=3, validate="key")
    # en_parallel ['validatecommand'] = (en_x_max.register(en_x_val), '%P', '%d')
    en_parallel.pack(side=LEFT)
    en_parallel.insert(END, '10')

    lbl = Label(frame_fw_controls, text=f'{" ":2}Channel =')
    lbl.pack(side=LEFT)

    combo_group_channel = Combobox(frame_fw_controls, state='readonly', width=8)
    combo_group_channel['values'] = ['auto', 'gsm', 'wifi', 'ethernet']
    combo_group_channel.pack(side=LEFT, anchor=NW, padx=3, pady=3)
    combo_group_channel.set('auto')

    lbl = Label(frame_fw_controls, text=f'{" ":20}')
    lbl.pack(side=LEFT)

    combo_ftp = Combobox(frame_fw_controls, state='normal', width=25)
    combo_ftp.pack(side=LEFT, anchor=NW, padx=3, pady=3)
    combo_ftp['values'] = 'cityair.uniscan.biz:6658'
    combo_ftp.set('cityair.uniscan.biz:6658')

    btn_ftp_refresh = Button(frame_fw_controls, text="Refresh", command=ftp_refresh)
    btn_ftp_refresh.pack(side=LEFT, padx=3)

    combo_fw_list = Combobox(frame_fw_controls, state='readonly', width=25, height=50)
    combo_fw_list.pack(side=LEFT, anchor=NW, padx=3, pady=3)
    combo_fw_list.bind('<<ComboboxSelected>>', combo_fw_list_select_event)

    combo_fw_list2 = Combobox(frame_fw_controls, state='disable', width=25, height=50)
    combo_fw_list2.bind('<<ComboboxSelected>>', combo_fw_list2_select_event)
    combo_fw_list2.pack(side=LEFT, anchor=NW, padx=3, pady=3)

    btn_fw_start = Button(frame_fw_controls, state='normal', text="Start", command=group_toggle)
    btn_fw_start.pack(side=LEFT, padx=3)

    frame_fw_log = Frame(tab_fw_update)
    frame_fw_log.pack(side=TOP, anchor=NW, padx=0, pady=0, fill=X, expand=1)

    log_fw = ScrolledText(master=frame_fw_log, width=300, height=100, wrap=WORD)
    log_fw.insert(END, '')
    log_fw.configure(state='disabled')
    log_fw.bind('<MouseWheel>', log_fw_wheel_event)
    log_fw.pack(side=RIGHT)
    log_fw_win = Toplevel(log_fw)
    log_fw_win.bind('<MouseWheel>', log_fw_wheel_event)
    log_fw_win.overrideredirect(True)
    # log_fw_win.maxsize(width=400, height=0)
    log_fw_win_text = Text(log_fw_win, padx=10, pady=10, wrap=WORD, borderwidth=0, bg='grey95',
                           highlightbackground="grey", highlightthickness=1)
    log_fw_win_text.pack()
    log_fw_win_y = 0
    log_fw_win_fix = False

    ftp_refresh()
    write_to_file_init()
    thread_clients_cleanup = threading.Thread(target=clients_cleanup, daemon=True)
    thread_clients_cleanup.start()
    thread_clients_check = threading.Thread(target=clients_check, daemon=True)
    thread_clients_check.start()

    # window.state('zoomed')
    window.mainloop()
