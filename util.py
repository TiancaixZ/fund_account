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


class show_Log:
    """
    实现log输出
    """

    def __init__(self, widget):
        self._widget = widget

    def show(self, content):
        self._widget.insert('end', content + '\n')
        self._widget.see('end')
