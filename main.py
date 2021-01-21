# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

# Press the green button in the gutter to run the script.
from read import read_list
from write import sheet

if __name__ == '__main__':
    # """提示用户输入文件地址"""
    # address = input("请输入文件地址：")
    # """提示用户输入月份"""
    # month = input("请输入月份：")
    # """提示用户输入文件输入地址"""
    # path = input("请输入文件输入地址：")

    # create_sheet(month, sheet_list(month, get_wb(address)), path)
    summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list = \
        read_list("2101", "/Users/chenguozhen/Downloads/2021年资金账（模板0.xlsx")
    sheet("2101", summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_list,
          "/Users/chenguozhen/Downloads/")