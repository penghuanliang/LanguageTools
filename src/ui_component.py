import sys, os
from tkinter import Button, Label, Entry, W, CENTER, RAISED, StringVar, Frame
from tkinter.filedialog import askdirectory
from win10toast import ToastNotifier
from pathlib import Path

from generate_xsls import NoTranslate


class UINoTranslate:
    # 按钮样式
    btn_style = {
        'text': '请选择工程下的res目录',
        'font': ('Arial', 12),
        'bg': '#28468D',
        'fg': 'white',
        'width': 6,
        'height': 2,
        'justify': CENTER,
        'relief': RAISED,
        'bd': 0,
        'activebackground': '#263C66',
        'activeforeground': 'white'
    }

    callback = None

    def set_callback(self, cb):
        self.callback = cb

    def select_res_file(self):
        in_options = {
            'initialdir': 'E:/as_workspace/Android/FXProject/src/main/res',
            'title': '选择一个res目录',
            'mustexist': False
        }

        _in = askdirectory(**in_options)
        self.e1.set(_in)

    def select_dir(self):
        out_options = {
            'initialdir': 'E:/translate',
            'title': '选择一个目录用于存放未翻译的 xlsx 文件',
            'mustexist': False
        }
        _out = askdirectory(**out_options)
        self.e2.set(_out)

    def show_msg(self, message):
        ico = self.resource_path(os.path.join('.\\icon\\favicon.ico'))
        toast = ToastNotifier()
        toast.show_toast(title='导出提示', msg=message, icon_path=ico,
                         duration=3, threaded=True)

    def run(self):
        e1_path = Path(self.e1.get())
        e2_path = Path(self.e2.get())
        if e1_path.exists() is False or e2_path.exists() is False or self.e1.get() == '' or self.e2.get() == '':
            self.show_msg("文件不存在！请检查文件目录是否存在！")
            return

        NoTranslate().generate_xsl(self.e1.get(), self.e2.get())
        self.show_msg("所有文本均已导出完毕！")

    # 生成资源文件目录访问路径
    def resource_path(self, relative_path):
        if getattr(sys, 'frozen', False):  # 是否Bundle Resource
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def back(self):
        self.frame.destroy()
        if self.callback is None:
            return
        self.callback()

    def __init__(self, parent):
        self.win = parent

        frame = Frame()

        self.frame = frame

        Label(master=frame, text='选择一个res目录', font=('Arial', 14)).grid(row=0, column=0, sticky=W, padx=15, pady=15)
        Label(master=frame, text='选择一个将要输出的目录', font=('Arial', 14)).grid(row=2, column=0, sticky=W, padx=15)

        self.e1 = StringVar()
        self.e2 = StringVar()

        entry1 = Entry(master=frame, textvariable=self.e1, bd=0, font=('Arial', 14))
        entry2 = Entry(master=frame, textvariable=self.e2, bd=0, font=('Arial', 14))

        entry1.grid(row=1, column=0, sticky=W, padx=15, pady=10, ipadx=230, ipady=14)
        entry2.grid(row=3, column=0, sticky=W, padx=15, pady=10, ipadx=230, ipady=14)

        Button(master=frame, text='打开', command=self.select_res_file, cnf=self.btn_style).grid(row=1, column=1,
                                                                                               sticky=W,
                                                                                               padx=10,
                                                                                               pady=10)
        Button(master=frame, text='打开', command=self.select_dir, cnf=self.btn_style).grid(row=3, column=1, sticky=W,
                                                                                          padx=10, pady=10)

        Label(master=frame).grid(row=4, column=0, pady=20)

        Button(master=frame, text='执行', command=self.run, cnf=self.btn_style).grid(row=5, column=0, sticky=W, padx=10,
                                                                                   pady=10,
                                                                                   columnspan=3, ipadx='360')
        # tuple callback
        Button(master=frame, text='返回', command=self.back,
               cnf=self.btn_style).grid(row=6, column=0, sticky=W,
                                        padx=10,
                                        pady=0,
                                        columnspan=3, ipadx='360')
        frame.pack()
