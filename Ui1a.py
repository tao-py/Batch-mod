from Funcs1a import *
import tkinter as tk

class MyGui:
    def __init__(self, window, name, mingling, row_col, *sticky):
        """
        __init__是构造函数。它创建类对象是被调用，用于初始化对象的属性
        :param window: 窗口
        :param name: 按钮name
        :param mingling: 命令对应的函数
        :param row_col: 行列值,元组或列表（x，y）/[x,y]
        """
        self.window = window
        self.TXT = name
        self.mingling = mingling
        self.row_col=row_col
        self.sticky = sticky
        self.button = tk.Button(self.window, text="{}".format(self.TXT), command=self.mingling)
        self.button = tk.Button(self.window, text="{}".format(self.TXT), command=self.mingling)
        self.button.grid(row=row_col[0], column=row_col[1],sticky=self.sticky)
class GUI:
    def __init__(self, window):
        self.window = window
    def my_label(self,name,row_col):
        self.TXT = name
        self.row_col = row_col
        self.print_label = tk.Label(self.window, text="{}".format(self.TXT))
        self.print_label.grid(row=row_col[0], column=row_col[1])

def open_new_window1():
    new_window = tk.Toplevel()
    new_window.title("用例格式检测")
    new_window.geometry(centre_window(300, 100, new_window))

    button1 = tk.Button(new_window, text="实测值宽度调整",command=run_function1)
    button1.pack()


    button2 = tk.Button(new_window, text="序列号调整")
    button2.pack()
def open_new_window2():
    new_window = tk.Toplevel()
    new_window.title("同类用例自动生成")
    new_window.geometry(centre_window(300, 100, new_window))

    button1 = tk.Button(new_window, text="待开发1")
    button1.pack()

    button2 = tk.Button(new_window, text="待开发2")
    button2.pack()
