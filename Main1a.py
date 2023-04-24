from Ui1a import *
import tkinter as tk
from Funcs1a import *

''' print与GUI关联 '''
class Print_GUI:
    def __init__(self,root):

        self.root = root

        # 创建一个文本框，并将其放置在窗口左侧
        self.textbox = tk.Text(self.root, wrap="word",height=10, width=50)
        self.textbox.grid(row=8, column=1)

        # 重定向print函数的输出，将其显示在文本框中
        self.stdout = StdoutRedirector(self.textbox)

    def run(self):
        self.root.mainloop()

class StdoutRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

        # 保存原始的sys.stdout对象
        self.old_stdout = None

        # 重定向sys.stdout到self.write方法
        self.activate()

    def activate(self):
        import sys
        self.old_stdout = sys.stdout
        sys.stdout = self

    def write(self, text):
        # 将输出文本显示在文本框中，并将光标移动到文本末尾
        self.textbox.insert("end", text)
        self.textbox.see("end")
        self.textbox.update_idletasks()

        # 将输出文本也打印到控制台，以便调试
        self.old_stdout.write(text)

    def flush(self):
        pass

# 创建GUI对象，并运行程序
def main():
    root = tk.Tk()
    root.title("用例Aide1.0")

    gui = Print_GUI(root)
    root.geometry(centre_window(500, 400,root))
    GUI_text1=GUI_text(root)
    GUI1=GUI(root)
    GUI1.my_label("Excel文件\n批量替换：",(0,0))
    GUI1.my_label("替换前内容:",(2,1))
    GUI1.my_label("替换后内容:", (4, 1))
    GUI1.my_label("运行显示:", (6, 0))
    GUI1.my_label("用例工具：", (9, 0))

    MyGui(root, "选择文件", select_files, (0, 1))
    MyGui(root, "确定&Add", GUI_text1.get_text, (5,0))
    MyGui(root, "开始替换", begin_run, (6, 1))

    MyGui(root, "用例格式检测", open_new_window1, (10, 1),"w")
    MyGui(root, "同类用例自动生成", open_new_window2, (10, 1),"e")
    # root.button = tk.Button(root, text="", command=select_files)
    # root.button.grid(row=1, column=0,sticky="w")
    gui.run()


if __name__ == '__main__':
    main()





