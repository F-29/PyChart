from tkinter import *
import time
import gaugelib
import datetime
import utils

# region static for recurrent function
g_value = 0
x = 0
store_data = []

# endregion

date = str(str(datetime.datetime.now()).replace('.', "-"))
date = str(str(date).replace(':', "-"))
file_path = str(utils.resource_path("data") + "_" + date + ".xlsx").replace(' ', "_")


def read_serial_and_save():
    global x
    g_value = utils.read_arduino_serial()
    global store_data
    store_data += [g_value]
    p1.set_value(int(g_value))
    x += 1
    if x > 1000:
        #        graph1.draw_axes()
        x = 0
    win.after(1000, read_serial_and_save)
    utils.save_to_excel(store_data, str(file_path))


if __name__ == '__main__':
    while True:
        if utils.is_arduino_connected():
            win = utils.create_win()
            port = utils.find_arduino_connected_port()
            p1 = gaugelib.DrawGauge2(
                win,
                max_value=125.0,
                min_value=0.0,
                size=435,
                bg_col='black',
                unit="ADC value %", bg_sel=2)
            p1.pack()
            read_serial_and_save()
            btn = Button(win, text="Exit", command=exit)
            btn.pack()
            mainloop()
        else:
            if utils.is_arduino_connected():
                break

            root = Tk()
            root.title("Arduino GUI interface")
            print("Arduino is NOT connected...")
            log = Text(root, width=28, height=2)
            log.insert(INSERT, "Arduino is NOT connected...")
            btn = Button(root, text="Exit", command=exit)
            log.pack()
            btn.pack()
            time.sleep(2.5)
            mainloop()
