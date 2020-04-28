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

# region Excel path fixing
file_path = utils.fix_path()


# endregion


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
            exit_btn = Button(win, text="Exit", command=exit)
            exit_btn.pack(side=BOTTOM)
            p1 = gaugelib.DrawGauge2(
                win,
                max_value=125.0,
                min_value=0.0,
                size=435,
                bg_col='black',
                unit="ADC value %", bg_sel=2)
            p1.pack(side=BOTTOM)
            read_from_excel_btn = Button(win, text="load excel file", command=utils.read_excel_and_play)
            read_from_excel_btn.pack(side=TOP)
            read_serial_and_save()
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
