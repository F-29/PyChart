import os
import datetime
import time
from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfile
import pandas as pd
import serial.tools.list_ports

g_value = 0
x = 0
numbers = 0


def all_children(window):
    _list = window.winfo_children()

    for item in _list:
        if item.winfo_children():
            _list.extend(item.winfo_children())

    return _list


def save_to_excel(data, path_to_save: str, columns=None) -> pd.DataFrame:
    """
    :param data:
    :param path_to_save:
    :param columns:
    :return:
    """
    if columns is None:
        columns = ["numbers"]
    excel_frame = pd.DataFrame(data, columns=columns)
    excel_frame.to_excel(path_to_save)

    return excel_frame


def find_arduino_connected_port():
    ls = [tuple(p) for p in list(serial.tools.list_ports.comports())]
    if len(ls):
        return ls[0][0]
    return []


def is_arduino_connected():
    if len(find_arduino_connected_port()) > 0:
        return True
    return False


def read_arduino_serial():
    port = find_arduino_connected_port()
    value = serial.Serial(port, 9600, timeout=None)
    value.setDTR(1)
    value = str(value.readline())
    value = value.replace("'", '')
    value = value.replace(" ", '')
    value = value.replace("b", '')
    value = value.replace("\\r", '')
    value = value.replace("\\n", '')
    return int(int(value) / 920.5 * 132)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def fix_path() -> str:
    date = str(str(datetime.datetime.now()).replace('.', "-"))
    date = str(str(date).replace(':', "-"))
    return str(resource_path("data") + "_" + date + ".xlsx").replace(' ', "_")


def create_win():
    win = tk.Tk()
    # a5 = PhotoImage(file=str(resource_path("g1.png")))
    # win.tk.call('wm', 'iconphoto', win._w, a5)
    win.title("Readings from arduino")
    win.geometry("1000x800+0+0")
    win.resizable(width=True, height=True)
    win.configure(bg='black')

    return win


def read_from_excel():
    file = askopenfilename()
    if len(file):
        return pd.read_excel(file).numbers.to_list()
    return False


def play_record(delay_ms: int = 1000):
    file_numbers = read_from_excel()
    for number in file_numbers:
        time.sleep(delay_ms)
        return number


def save_as_excel(data: list):
    """
    :param data:
    :return:
    """
    file_name = ("Excel Workbook", "*.xlsx")
    file = asksaveasfile(filetypes=[file_name], defaultextension=file_name)
    data_frame = pd.DataFrame(data, columns=['numbers'])
    data_frame.to_excel(file.name)
    return file
