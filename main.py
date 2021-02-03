# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

# Press the green button in the gutter to run the script.
from read import read_list
from write import sheet
from tkinter import *
import tkinter.filedialog


if __name__ == '__main__':
    # """提示用户输入文件地址"""
    # address = input("请输入文件地址：")
    # """提示用户输入月份"""
    # month = input("请输入月份：")
    # """提示用户输入文件输入地址"""
    # path = input("请输入文件输入地址：")

    # create_sheet(month, sheet_list(month, get_wb(address)), path)
    # summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list = \
    #     read_list("2101", "/Users/chenguozhen/Downloads/2021年资金账（模板0.xlsx")
    # sheet("2101", summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list,
    #       "/Users/chenguozhen/Downloads/")


    def selectPath():
        """
        打开文件
        :return:
        """
        open_path = tkinter.filedialog.askopenfilename()
        path.set(open_path)

        return open_path

    def selectsavePath():
        """
        选择文件保存地址
        :return:
        """
        path_save = tkinter.filedialog.askdirectory()
        save_path.set(path_save)

        return path_save

    def run():
        month = month_entry.get()

        summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list = \
            read_list(month, selectPath())
        sheet(month, summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list,
              selectsavePath()+"/")

    window = Tk()
    path = StringVar()
    save_path = StringVar()
    window.title("南京朗坤医药有限责任公司")
    window.geometry('500x300')

    Label(window, text="选择文件:", font=('Arial', 14), width=10, heigh=2).grid(row=0, column=0)

    Entry(window, textvariable=path).grid(row=0, column=1)

    Button(window, text='选择文件', command=selectPath).grid(row=0, column=2)

    Label(window, text="输入密码：", font=('Arial', 14), width=10, heigh=2).grid(row=1, column=0)

    Entry(window, show='*', font=('Arial', 14)).grid(row=1, column=1)

    Label(window, text="输入月份：", font=('Arial', 14), width=10, heigh=2).grid(row=2, column=0)

    month_entry = Entry(window, show=None, font=('Arial', 14))
    month_entry.grid(row=2, column=1)

    Label(window, text="选择保存文件地址:", font=('Arial', 14), width=20, heigh=2).grid(row=3, column=0)

    Entry(window, textvariable=save_path).grid(row=3, column=1)

    Button(window, text='选择文件夹', command=selectsavePath).grid(row=3, column=2)

    Button(window, text='ok', command=run, font=('Arial', 12), width=20, height=2).grid(row=4, column=0)

    window.mainloop()


