import os, sys
from tkinter import Tk, Frame, Button, CENTER, RAISED, W, Label
from ui_component import UINoTranslate
from ui_xlsx_to_xml import UiXLSXToXML
from ui_xml_to_xsls import UiXMLToXLSX

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


# 生成资源文件目录访问路径
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # 是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def home_page():
    root.withdraw()
    one = Frame(root)
    Label(master=one).grid(row=0, pady=70)
    Button(master=one, text='Android工具：XLSX  to  XML', cnf=btn_style, command=lambda: {
        one.destroy(),
        xlsx_to_xml()
    }).grid(row=1, sticky=W, padx=10, pady=10,
            ipadx='360')
    Button(master=one, text='Android工具： XML  to XLSX', cnf=btn_style, command=lambda: {
        one.destroy(),
        xml_to_xlsx()
    }).grid(row=2, sticky=W, padx=10, pady=10,
            ipadx='360')
    Button(master=one, text='Android工具：导出未翻译文本', cnf=btn_style, command=lambda: {
        one.destroy(),
        out_no_translate()
    }).grid(row=3, sticky=W, padx=10,
            pady=10, ipadx='360')
    one.pack()

    root.update()
    root.deiconify()


def xlsx_to_xml():
    root.withdraw()
    # 将要开发的内容
    _ui = UiXLSXToXML(root)
    # tuple
    _ui.set_callback(lambda: {
        home_page()
    })
    root.update()
    root.deiconify()


def xml_to_xlsx():
    root.withdraw()
    # 将要开发的内容
    _ui = UiXMLToXLSX(root)
    # tuple
    _ui.set_callback(lambda: {
        home_page()
    })
    root.update()
    root.deiconify()


def out_no_translate():
    root.withdraw()
    # 将要开发的内容
    _ui = UINoTranslate(root)
    # tuple
    _ui.set_callback(lambda: {
        home_page()
    })
    root.update()
    root.deiconify()


if __name__ == '__main__':
    root = Tk()
    ico = resource_path(os.path.join('.\\icon\\favicon.ico'))
    root.title('自动化处理未翻译文本')
    root.iconbitmap(ico)
    root.resizable(width=False, height=False)
    w = 800
    h = 500
    left = int((root.winfo_screenwidth() - w) / 2)
    top = int((root.winfo_screenheight() - h - 30) / 2)
    root.geometry(f'{w}x{h}+{left}+{top}')

    home_page()

    root.mainloop()
