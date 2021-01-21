from abc import ABCMeta


class data(object, metaclass=ABCMeta):
    """银行卡"""

    def __init__(self, name):
        """
        初始化方法
        :param name:
        """
        self._name = name

    @property
    def name(self):
        return self._name


class Summary_data(data):

    def __init__(self, name, lmb, income, expenditure, contacts):
        """
        上月余额，收入，支出，往来，余额
        :param name:
        :param lmb:
        :param income:
        :param expenditure:
        :param contacts:
        """
        super().__init__(name)
        self._lmb = lmb
        self._income = income
        self._expenditure = expenditure
        self._contacts = contacts

    @property
    def lmb(self):
        return self._lmb

    @property
    def income(self):
        return self._income

    @property
    def expenditure(self):
        return self._expenditure

    @property
    def contacts(self):
        return self._contacts

    @property
    def balance(self):
        """
        计算余额
        :return:
        """
        if self._lmb is not None and self._income is not None and \
                self._contacts is not None and self._expenditure is not None and \
                self._name is not None:
            return float(self._lmb) + float(self._income) - \
                   float(self._expenditure) + float(self._contacts)
        else:
            return -1


class Inc_Exp_data(data):

    def __init__(self, name, pri_com, inc_exp):
        """
        单位名称/负责人 收入 支出
        :param name:
        :param pri_com:
        :param inc_exp:
        """
        super().__init__(name)
        self._pri_com = pri_com
        self._inc_exp = inc_exp

    @property
    def pri_com(self):
        return self._pri_com

    @property
    def inc_exp(self):
        return self._inc_exp
