# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : pandas_financial.py
# @Software : PyCharm

import operator
import pandas as pd
import ask_helper
import matplotlib.pyplot as plt

pd.set_option('display.max_columns', None)


def get_normal_code():
    codes = ask_helper.read_excel("Acode.xls")
    print(codes)
    return codes


def get_normal_data():
    codes = get_normal_code()
    for dict in codes:
        for item in dict.items():
            pf = pandas_finance()
            pf.set_normal_code(item[0], item[1])
            pf.set_read_file_name()
            tp_profit = pf.get_excel_profit_data()
            tp_cash = pf.get_excel_cashflow_data()
            print(item[1])
            print(tp_profit)
            print(tp_cash)
            pfd = pd.Series(tp_profit)
            chd = pd.Series(tp_cash)
            d = {'profit': pfd, 'cash': chd}
            df = pd.DataFrame(d)
            df.plot()
            plt.rcParams['font.sans-serif'] = ['KaiTi']
            plt.title(item[1])
            plt.savefig("./Statistic/{}_{}.png".format(item[0], item[1]))
            plt.close()
            # plt.show()


# 一个code一个pandasFinance实例
class pandas_finance:
    @classmethod
    def is_number(cls, str):
        try:
            if str == 'NaN':
                return 0
            float(str)
            return float(str)
        except ValueError:
            return 0

    @classmethod
    def get_missed_columns_index(cls, list1, list2):
        """
        :param list1: 比较全的list
        :param list2: 不全的list
        :return: 第几个不全，d得到list1中的下标  数组
        """
        result = []
        for index, item in enumerate(list1):
            if item in list2:
                continue
            else:
                result.append(index)
        return result

    def __init__(self):
        self.code = ""
        self.name = ""
        self.excel_file_financial = "./Excel/statistic_{}_{}_financial.xls"
        self.excel_file_profit = "./Excel/statistic_{}_{}_profit.xls"
        self.excel_file_cashflow = "./Excel/statistic_{}_{}_cashflow.xls"
        self.cash_data = []
        self.profit_data = []
        self.financial_data = []

    def set_normal_code(self, code, name):
        self.code = code
        self.name = name

    def set_read_file_name(self):
        self.excel_file_financial = self.excel_file_financial.format(self.code, self.name)
        self.excel_file_profit = self.excel_file_profit.format(self.code, self.name)
        self.excel_file_cashflow = self.excel_file_cashflow.format(self.code, self.name)

    def get_excel_finance_data(self):
        pro_data = pd.read_excel(self.excel_file_financial)
        profit = pro_data.iloc[1][1:]  # Series数据 带一个列名的行数据  且过滤第一列，它不是数据
        grouped = profit.groupby(lambda a: str(a)[:4])
        result = []
        pfd = {}
        for key, values in grouped:
            values = values.apply(lambda a: self.is_number(a))
            values = values.sort_index(ascending=False)
            ln = len(values)
            i = 0
            for ky, item in values.items():
                if i < (ln - 1):
                    nt = values[i + 1]
                    cur = values[i]
                    pfd.update({ky: cur - nt})
                if i == (ln - 1):
                    cur = values[i]
                    pfd.update({ky: cur})
                i = i + 1
        pfd = sorted(pfd.items(), key=lambda d: d[0], reverse=False)
        # result = sorted(result, key=operator.itemgetter('date'), reverse=False)
        return dict(pfd)

    def get_excel_profit_data(self):
        pro_data = pd.read_excel(self.excel_file_profit)
        profit = pro_data.iloc[1][1:]  # Series数据 带一个列名的行数据  且过滤第一列，它不是数据
        grouped = profit.groupby(lambda a: str(a)[:4])
        result = []
        pfd = {}
        for key, values in grouped:
            values = values.apply(lambda a: self.is_number(a))
            values = values.sort_index(ascending=False)
            ln = len(values)
            i = 0
            for ky, item in values.items():
                if i < (ln - 1):
                    nt = values[i + 1]
                    cur = values[i]
                    pfd.update({ky: cur - nt})
                if i == (ln - 1):
                    cur = values[i]
                    pfd.update({ky: cur})
                i = i + 1
        pfd = sorted(pfd.items(), key=lambda d: d[0], reverse=False)
        # result = sorted(result, key=operator.itemgetter('date'), reverse=False)
        return dict(pfd)

    def get_excel_cashflow_data(self):
        cash_data = pd.read_excel(self.excel_file_cashflow)
        # cash = cash_data.iloc[[4]]
        cash = cash_data.iloc[4][1:]  # Series数据 带一个列名的行数据  且过滤第一列，它不是数据
        grouped = cash.groupby(lambda a: str(a)[:4])  # 把index，也就是年，拿出来并截取年，按年分组
        result = []
        cfd = {}
        # grouped.size().index # 分组后的索引，也就是key列表
        for key, values in grouped:  # 可以直接拿到key和对应的数据列表
            values = values.apply(lambda a: self.is_number(a))  # 转成float
            values = values.sort_index(ascending=False)  # 倒序
            # print(values.map(type)) # 打印所含类型
            ln = len(values)
            i = 0
            for ky, item in values.items():
                if i < (ln - 1):
                    nt = values[i + 1]
                    cur = values[i]
                    result.append({'date': ky, 'cash': cur - nt})
                    cfd.update({ky: cur - nt})
                if i == (ln - 1):
                    cur = values[i]
                    result.append({'date': ky, 'cash': cur})
                    cfd.update({ky: cur})
                i = i + 1
        cfd = sorted(cfd.items(), key=lambda a: a[0], reverse=False)
        # result = sorted(result, key=operator.itemgetter('date'), reverse=False)
        return dict(cfd)

    def get_excel_cashflow_all(self):
        cash_data = pd.read_excel(self.excel_file_cashflow, header=0, index_col=0)
        cash_data = cash_data.T

        filterCol = ['经营活动现金流入小计', '经营活动现金流出小计', '经营活动产生的现金流量净额',
                     '销售商品、提供劳务收到的现金', '购买商品、接受劳务支付的现金']

        cashflow = cash_data.filter(items=filterCol)
        # 按行index分组
        result = []
        grouped = cashflow.groupby(lambda a: str(a)[:4], axis=0)
        for ky, itemlist in grouped:  # ky 是年
            print(ky)
            itemresult = {}
            for key, item in itemlist.items():  # key 是项目名
                item = item.apply(lambda a: self.is_number(a))
                item = item.sort_index(ascending=True)  # 排序
                new = item.diff().fillna(item[0])  # 第一行是空的，NAN直接拿第一行数据
                itemresult.update({key: new})  # new 是Series
            print(itemresult)
            new_df = pd.DataFrame(itemresult).reset_index()  # 把index数据作为新增一列并用RangeIndex作为新的index
            new_df.insert(0, 'name', self.name)
            new_df.insert(0, 'code', self.code)
            data_list = [tuple(i) for i in new_df.values]
            result.append(data_list)
            # l =len(item)
            # i = 0
            # for ikey,it in item.items():
            #     print(key,ikey,it)
            #     if i<l-1:
            #         pass
            # item = item.apply(lambda a: self.is_number(a))
            # print(item)
            # for item in itemlist:
            #     print(item)
        return result
        # for col_name in cash_data:
        #     if col_name == "经营活动现金流入小计":
        #         print(type(cash_data[[col_name]]))
        # print(cash_data)
        # 获取列 没有列会报错！
        # cashflow = cash_data[['经营活动现金流入小计','经营活动现金流出小计', '经营活动产生的现金流量净额',
        #                       '销售商品、提供劳务收到的现金','购买商品、接受劳务支付的现金']]
        # print(cashflow)
        # cashflow = cash_data.loc[['经营活动现金流入小计', '经营活动现金流出小计', '经营活动产生的现金流量净额',
        #                           '销售商品、提供劳务收到的现金','购买商品、接受劳务支付的现金'
        #                           # '投资活动现金流入小计','投资活动现金流出小计','投资活动产生的现金流量净额',
        #                           # '筹资活动现金流入小计','筹资活动现金流出小计','筹资活动产生的现金流量净额'
        #                           ]]

        # return cashflow

    def get_excel_profit_all(self):
        profit_data = pd.read_excel(self.excel_file_profit, header=0, index_col=0)
        profit_data = profit_data.T
        filterCol = ['一、营业总收入',
                     '二、营业总成本', '销售费用', '管理费用', '研发费用', '财务费用', '投资收益',
                     '三、营业利润', '其中：非流动资产处置利得', '其中：非流动资产处置损失',
                     '五、净利润', '扣除非经常性损益后的净利润']
        profit = profit_data.filter(items=filterCol)
        result = []
        grouped = profit.groupby(lambda a: str(a)[:4], axis=0)
        for ky, itemlist in grouped:
            matched_cols = list(itemlist.columns.values)
            missed_cols = [(i, v) for i, v in enumerate(filterCol) if v not in matched_cols]
            print(missed_cols)
            # 补全缺失列
            for tup in missed_cols:
                to_add_name = filterCol[tup[0]]
                # print(self.code,tup[0],tup[1],to_add_name)
                # print("插入前：",tup)
                # print(itemlist)
                if to_add_name != tup[1]:
                    print("数据错误：", self.code, self.name, tup, 'findname:' + to_add_name)
                itemlist.insert(tup[0], to_add_name, 0)
                # print("插入后：",tup)
                # print(itemlist)

            # 处理数据：累减
            item_result = {}
            for key, item in itemlist.items():  # key 是项目名
                item = item.apply(lambda a: self.is_number(a))
                item = item.sort_index(ascending=True)  # 排序
                new = item.diff().fillna(item[0])  # 第一行是空的，NAN直接拿第一行数据
                item_result.update({key: new})  # new 是Series

            # 转换
            new_df = pd.DataFrame(item_result).reset_index()  # 把index数据作为新增一列并用RangeIndex作为新的index
            new_df.insert(0, 'name', self.name)
            new_df.insert(0, 'code', self.code)
            data_list = [tuple(i) for i in new_df.values]
            result.append(data_list)
        return result

    def get_excel_finical_all(self):
        financial_data = pd.read_excel(self.excel_file_financial, header=0, index_col=0)
        financial_data = financial_data.T
        filterCol = ['货币资金','应收票据及应收账款','预付款项','存货','流动资产合计',
                     '固定资产合计','在建工程合计', '无形资产', '商誉','长期待摊费用', '非流动资产合计',
                     '应付票据及应付账款', '流动负债合计', '非流动负债合计',
                     '实收资本（或股本）', '资本公积', '盈余公积', '未分配利润']
        financial = financial_data.filter(items=filterCol)
        result = []
        grouped = financial.groupby(lambda a: str(a)[:4], axis=0)

        for ky, itemlist in grouped:
            matched_cols = list(itemlist.columns.values)
            missed_cols = [(i, v) for i, v in enumerate(filterCol) if v not in matched_cols]
            print("缺失列：", missed_cols)
            # 补全缺失列
            for tup in missed_cols:
                to_add_name = filterCol[tup[0]]
                if to_add_name != tup[1]:
                    print("数据错误：", self.code, self.name, tup, 'findname:' + to_add_name)
                itemlist.insert(tup[0], to_add_name, 0)

            # 处理数据：累减
            item_result = {}
            for key, item in itemlist.items():  # key 是项目名
                item = item.apply(lambda a: self.is_number(a))
                item = item.sort_index(ascending=True)  # 排序
                #new = item.diff().fillna(item[0])  # 第一行是空的，NAN直接拿第一行数据
                item_result.update({key: item})  # new 是Series

            # 转换
            new_df = pd.DataFrame(item_result).reset_index()  # 把index数据作为新增一列并用RangeIndex作为新的index
            new_df.insert(0, 'name', self.name)
            new_df.insert(0, 'code', self.code)
            data_list = [tuple(i) for i in new_df.values]
            result.append(data_list)
        print(result)
        return result
