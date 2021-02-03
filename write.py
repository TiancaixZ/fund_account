from openpyxl import Workbook
from openpyxl.styles import Side, Border, Alignment

from util import if_null


def sheet(date, summary_list, pri_inc_list, pri_exp_list, com_inc_list, com_exp_lis, path):
    month = str(date)[2:4]

    """创建sheet"""
    wb = Workbook()
    ws_summary = wb.create_sheet("汇总", 0)
    ws_income_principal = wb.create_sheet("收入_负责人", 1)
    ws_expenditure_principal = wb.create_sheet("支出_负责人", 2)
    ws_income_company = wb.create_sheet("收入_公司", 3)
    ws_expenditure_company = wb.create_sheet("支出_公司", 4)

    summary_sheet(ws_summary, month, summary_list)

    for item in pri_inc_list:
        print(item.pri_com + str(len(pri_inc_list)))

    incom_exp_sheet(ws_income_principal, month, pri_inc_list, 0)
    incom_exp_sheet(ws_expenditure_principal, month, pri_exp_list, 1)
    incom_exp_sheet(ws_income_company, month, com_inc_list, 2)
    incom_exp_sheet(ws_expenditure_company, month, com_exp_lis, 3)

    """文件保存"""
    path = path + "test.xlsx"
    wb.save(path)
    print("文件保存成功..............................OK")


def muban_list(repat_list):
    """
    表头名 （卡名）  去重
    :param repat_list:
    :return:
    """
    name_old_list = []
    pricom_old_list = []
    for item in repat_list:
        name_old_list.append(item.name)
        pricom_old_list.append(item.pri_com)
    name_new_list = list(set(name_old_list))
    pricom_new_list = list(set(pricom_old_list))
    return name_new_list, pricom_new_list


def muban(name_list, pricom_list, ws):
    """
    模板生成 表头 首列
    :param name_list:
    :param pricom_list:
    :param ws:
    :return:
    """
    ws["A1"].border = sheet_border()
    ws["A1"].alignment = sheet_alignment()
    for column in range(2, len(name_list)+1):
        _ = ws.cell(column=column, row=2, value=name_list[column-1])    # 卡名
    for row in range(3, len(pricom_list)+4):
        if row == len(pricom_list)+3:
            _ = ws.cell(column=1, row=row, value="合计")
        else:
            _ = ws.cell(column=1, row=row, value=pricom_list[row-3])  # 单位名称/负责人

    for col in range(1, len(name_list)+1):
        for row in range(2, len(pricom_list)+4):
            ws.cell(row, col).border = sheet_border()
            ws.cell(row, col).alignment = sheet_alignment()
    ws.column_dimensions['A'].width = 40  # 设置列宽40
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(name_list))  # 合并单元表头
    print("模板创建..............................OK")


def summary_sheet(ws, month, list):
    """
    汇总
    :param ws:
    :param month:
    :param list:
    :return:
    """
    ws["A1"] = "2021年" + month + "月份朗坤资金流向"
    ws["A1"].border = sheet_border()
    ws["A1"].alignment = sheet_alignment()
    ws["A2"] = "科目"
    ws["B2"] = "上月余额"
    ws["C2"] = "收入"
    ws["D2"] = "支出"
    ws["E2"] = "往来"
    ws["F2"] = "余额"
    ws.merge_cells('A1:F1')
    for col in ws['A2:F2']:
        for cell in col:
            cell.border = sheet_border()
            cell.alignment = sheet_alignment()
    ws.column_dimensions['A'].width = 20  # 设置列宽20
    print("汇总 表头创建..............................OK")

    total_lmb = 0
    total_expenditure = 0
    total_income = 0
    total_contacts = 0
    total_balance = 0
    for row in range(3, len(list) + 4):
        if row == len(list)+3:
            _ = ws.cell(row, 1, "合计")
            _ = ws.cell(row, 2, total_lmb)
            _ = ws.cell(row, 3, total_income)
            _ = ws.cell(row, 4, total_expenditure)
            _ = ws.cell(row, 5, total_contacts)
            _ = ws.cell(row, 6, total_balance)
        else:
            _ = ws.cell(row, 1, list[row - 3].name)
            _ = ws.cell(row, 2, list[row - 3].lmb)
            total_lmb = total_lmb + if_null(list[row - 3].lmb)
            _ = ws.cell(row, 3, list[row - 3].income)
            total_income = total_income + list[row - 3].income
            _ = ws.cell(row, 4, list[row - 3].expenditure)
            total_expenditure = total_expenditure + list[row - 3].expenditure
            _ = ws.cell(row, 5, list[row - 3].contacts)
            total_contacts = total_contacts + list[row - 3].contacts
            _ = ws.cell(row, 6, list[row - 3].balance)
            total_balance = total_balance + list[row - 3].balance
        for col in range(1, 7):
            _ = ws.cell(row, col).border = sheet_border()
            _ = ws.cell(row, col).alignment = sheet_alignment()

    print("汇总 数据填写..............................OK")


def incom_exp_sheet(ws, month, ready_list, pri_com_inc_exp):
    """
    公司/负责人 收入/支出 表
    :param ws:
    :param month:
    :param ready_list:
    :param pri_com_inc_exp:
    :return:
    """
    if pri_com_inc_exp == 0:
        var = "负责人"
        var_title = "收入"
    elif pri_com_inc_exp == 1:
        var = "负责人"
        var_title = "支出"
    elif pri_com_inc_exp == 2:
        var = "单位名称"
        var_title = "收入"
    elif pri_com_inc_exp == 3:
        var = "单位名称"
        var_title = "支出 "

    ws["A1"] = "202年" + month + "月份朗坤资金" + var_title
    ws["A2"] = var
    name_list, pricom_list = muban_list(ready_list)
    muban(name_list, pricom_list, ws)

    for item in ready_list:
        for row in range(3, len(pricom_list) + 4):
            for column in range(2, len(name_list) + 1):
                if item.name == name_list[column-1] and item.pri_com == pricom_list[row-4]:
                    _ = ws.cell(column=column, row=row, value=item.inc_exp)


def sheet_border():
    """
    边框
    :return:
    """
    thin = Side(border_style="thin", color="000000")  # 边框样式，颜色
    border = Border(left=thin, right=thin, top=thin, bottom=thin)  # 边框的位置

    return border


def sheet_alignment():
    """
    对齐
    :return:
    """
    align = Alignment(horizontal='center', vertical='center')  # 剧中对齐
    return align
