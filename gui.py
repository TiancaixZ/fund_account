import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.scrolledtext

from read import Read
from util import show_Log
from write import Write


class Fund(tk.Tk):
    def __init__(self):
        super().__init__()

        self.path = tk.StringVar()
        self.save_path = tk.StringVar()
        self.scr_text = tk.StringVar()

        self.title("南京朗坤医药有限责任公司")
        self.geometry("500x400")

        tk.Label(self, text="选择文件:", font=('Arial', 14), width=10, heigh=2) \
            .grid(row=0, column=0)

        self.pathEntry = tk.Entry(self, textvariable=self.path)
        self.pathEntry.grid(row=0, column=1)

        tk.Button(self, text='选择文件', command=self.selectPath).grid(row=0, column=2)

        tk.Label(self, text="输入密码：", font=('Arial', 14), width=10, heigh=2) \
            .grid(row=1, column=0)

        self.passwordEntry = tk.Entry(self, show='*', font=('Arial', 14))
        self.passwordEntry.grid(row=1, column=1)

        tk.Label(self, text="输入月份：", font=('Arial', 14), width=10, heigh=2) \
            .grid(row=2, column=0)

        self.month_entry = tk.Entry(self, show=None, font=('Arial', 14))
        self.month_entry.grid(row=2, column=1)

        tk.Label(self, text="选择保存文件地址:", font=('Arial', 14), width=20, heigh=2) \
            .grid(row=3, column=0)

        self.save_pathEntry = tk.Entry(self, textvariable=self.save_path)
        self.save_pathEntry.grid(row=3, column=1)

        tk.Button(self, text='选择文件夹', command=self.selectsavePath) \
            .grid(row=3, column=2)

        tk.Button(self, text='ok', command=self.run, font=('Arial', 12), width=20, height=2) \
            .grid(row=4, column=0)

        Information = tk.LabelFrame(self, text="操作信息", padx=10, pady=10)
        Information.place(x=10, y=200)
        self.scr = tk.scrolledtext.ScrolledText(Information, width=50, heigh=2, font=('Arial', 14), wrap=tk.WORD)
        self.scr.grid()

    def selectPath(self):
        """
        打开文件
        :return:
        """
        open_path = tkinter.filedialog.askopenfilename()
        self.path.set(open_path)

    def selectsavePath(self):
        """
        选择文件保存地址
        :return:
        """
        path_save = tkinter.filedialog.askdirectory()
        self.save_path.set(path_save)

    def run(self):

        if self.passwordEntry.get() == '':
            tkinter.messagebox.showerror('错误', '密码没有输入')
        elif self.month_entry.get() == '':
            tkinter.messagebox.showerror('错误', '月份没有输入')
        elif self.month_entry.get() != '' and self.passwordEntry.get() != '':
            show = show_Log(self.scr)
            read = Read(self.passwordEntry.get(), self.pathEntry.get(), self.month_entry.get(), show)
            write = Write(self.month_entry.get(), self.save_pathEntry.get(), show)
            write.summary_sheet(read.summary_list)
            write.ws_income_principal_sheet(read.pri_inc_list)
            write.ws_expenditure_principal_sheet(read.pri_exp_list)
            write.ws_income_company_sheet(read.com_inc_list)
            write.ws_expenditure_company_sheet(read.com_exp_list)
            write.save()
            # test_list = read.com_exp_list
            # for item in test_list:
            #     print("test" + item.name + str(item.pri_com) + str(item.inc_exp))



