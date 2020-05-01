from tkinter import *
import time
import gaugelib
import utils

# region static for recurrent function
g_value = 0
x = 0
# z = 0
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
        x = 0
    win.after(50, read_serial_and_save)
    # utils.save_to_excel(store_data, str(file_path)) # for automatic save


# def chart_record():
#     global z
#     # number = utils.play_record(1000)
#     number = 5
#     new_win = utils.create_win()
#     new_exit_btn = Button(new_win, text="Exit", command=exit)
#     new_exit_btn.pack(side=BOTTOM)
#     p = gaugelib.DrawGauge2(
#         new_win,
#         max_value=125.0,
#         min_value=0.0,
#         size=435,
#         bg_col='white',
#         unit="ADC value %", bg_sel=2)
#     widget_list_2 = utils.all_children(win)
#     for i in widget_list_2:
#         i.pack_forget()
#     if number:
#         p.set_value(int(number))
#         z += 1
#         if z > 1000:
#             z = 0
#         win.after(100, read_serial_and_save)
#         p1.pack(side=BOTTOM)


if __name__ == '__main__':
    while True:
        if utils.is_arduino_connected():
            win = utils.create_win()
            port = utils.find_arduino_connected_port()
            # save_button =
            exit_btn = Button(win, text="Exit", command=exit)
            exit_btn.pack(side=BOTTOM)
            p1 = gaugelib.DrawGauge2(
                win,
                max_value=125.0,
                min_value=0.0,
                size=500,
                bg_col='black',
                unit="ADC value %", bg_sel=2)
            save_btn = Button(win, text="Save", command=lambda: utils.save_as_excel(store_data))
            save_btn.pack(side=TOP)
            p1.pack(side=BOTTOM, pady=20)
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
            # widget_list = utils.all_children(win)
            # for item in widget_list:
            #     item.pack_forget()
            mainloop()
