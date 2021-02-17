def if_null(value):
    """
    判断空
    :param value:
    :return:
    """
    var = 0
    if value is not None:
        var = value
    return var


def name_list(pri_inc_row, sheet):
    """
    单位/负责人  去重
    :param pri_inc_row:
    :param sheet:
    :return:
    """
    old_list = []
    if pri_inc_row == 0:
        for row in range(4, sheet.max_row + 1):
            if sheet.cell(row, 10).value is not None:
                old_list.append(sheet.cell(row, 10).value)  # 负责人
    else:
        for row in range(4, sheet.max_row + 1):
            if sheet.cell(row, 6).value is not None:
                old_list.append(sheet.cell(row, 6).value)  # 单位名称
    new_list = list(set(old_list))

    return new_list


def muban_list(repat_list):
    """
    表头名 （卡名）  去重
    :param repat_list:
    :return:
    """
    name_old_list = []
    pricom_old_list = []
    for item in repat_list:
        if item.pri_com != "往来":
            pricom_old_list.append(item.pri_com)
        name_old_list.append(item.name)
    name_new_list = list(set(name_old_list))
    pricom_new_list = list(set(pricom_old_list))
    return name_new_list, pricom_new_list


class show_Log:
    """
    实现log输出
    """

    def __init__(self, widget):
        self._widget = widget

    def show(self, content):
        self._widget.insert('end', content + '\n')
        self._widget.see('end')
