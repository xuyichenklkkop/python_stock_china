# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : ask_helper.py
# @Software : PyCharm

import requests
import re
import xlrd
import xlwt
from xlutils.copy import copy
from bs4 import BeautifulSoup
import threading



def read_excel(excelname):
    work_book = xlrd.open_workbook(excelname)
    one_sheet = work_book.sheets()[0]
    all_counts = one_sheet.nrows
    result = []
    for i in range(all_counts):
        if re.match("\*ST", one_sheet.row_values(i)[1]):
            continue
        result.append({one_sheet.row_values(i)[0]: one_sheet.row_values(i)[1]})
    return result

def get_financial_excel():
    codes = read_excel("ACode.xls")
    for code in codes:
        for key, val in code.items():
            askci = StockAskci()
            askci.set_file_ext("financial")
            askci.set_code(key, val)
            if (askci.check_file_not_exist()):
                askci.init_url()
                html = askci.get_financial_detail_data()
                trs = askci.parse_doc_data_return_trs(html)
                askci.open_excel_nomatter_exist(trs)
    print('financial end')

def get_profit_excel():
    codes = read_excel("ACode.xls")
    for code in codes:
        for key, val in code.items():
            askci = StockAskci()
            askci.set_file_ext("profit")
            askci.set_code(key, val)
            if (askci.check_file_not_exist()):
                askci.init_url()
                html = askci.get_profit_detail_data()
                trs = askci.parse_doc_data_return_trs(html)
                askci.open_excel_nomatter_exist(trs)
    print('profit_end')

def get_cashflow_excel():
    codes = read_excel("ACode.xls")
    for code in codes:
        for key, val in code.items():
            askci = StockAskci()
            askci.set_file_ext("cashflow")
            askci.set_code(key, val)
            if(askci.check_file_not_exist()):
                askci.init_url()
                html = askci.get_cashflow_detail_data()
                trs = askci.parse_doc_data_return_trs(html)
                askci.open_excel_nomatter_exist(trs)
    print('cashflow_end')




def get_three_excel():
    task_f = threading.Thread(target=get_financial_excel)
    task_c = threading.Thread(target=get_cashflow_excel)
    task_p = threading.Thread(target=get_profit_excel)
    task_f.start()
    task_c.start()
    task_p.start()


class StockAskci:

    def __init__(self):
        self.code_list = []
        self.financial_base_url = "https://s.askci.com/StockInfo/FinancialReport/BalanceSheet/?stockCode={}&theType=BalanceSheet&UnitName=%E4%B8%87%E5%85%83&dateRange=,5"  # 资产负债
        self.profit_base_url = "https://s.askci.com/StockInfo/FinancialReport/Profit/?stockCode={}&theType=Profit&UnitName=%E4%B8%87%E5%85%83&dateRange=,5"  # 利润表
        self.cashflow_base_url = "https://s.askci.com/StockInfo/FinancialReport/CashFlow/?stockCode={}&theType=CashFlow&UnitName=%E4%B8%87%E5%85%83&dateRange=,5"  # 现金流量表
        self.code = ""
        self.excelName = ".\Excel\statistic_{}_{}_{}.xls"
        self.excelExt = ""

    def set_file_ext(self, ext):
        self.excelExt = ext

    def set_code(self, _code, _name):
        self.code = _code
        self.excelName = self.excelName.format(self.code, _name, self.excelExt)

    def init_url(self):
        self.financial_base_url = self.financial_base_url.format(self.code)
        self.profit_base_url = self.profit_base_url.format(self.code)
        self.cashflow_base_url = self.cashflow_base_url.format(self.code)

    def check_file_not_exist(self):
        rst = True
        try:
            book = xlrd.open_workbook(self.excelName, formatting_info=True)
            rst = False
        except Exception as FileNotFoundError:
            pass
        finally:
            return rst


    def get_financial_detail_data(self):
        try:
            response = requests.get(self.financial_base_url, timeout=30000)
            response.encoding = "utf-8"
            html = response.text
            return html
        except Exception as e:
            print("http_err:", e)

    def get_profit_detail_data(self):
        try:
            response = requests.get(self.profit_base_url, timeout=30000)
            response.encoding = "utf-8"
            html = response.text
            return html
        except Exception as e:
            print("http_err:", e)

    def get_cashflow_detail_data(self):
        try:
            response = requests.get(self.cashflow_base_url, timeout=30000)
            response.encoding = "utf-8"
            html = response.text
            return html
        except Exception as e:
            print("http_err:", e)

    def parse_doc_data_return_trs(self, html):
        bs = BeautifulSoup(html, "html.parser")
        trs = bs.select("tr")
        return trs

    def open_excel_nomatter_exist(self, trs):
        try:
            book = xlrd.open_workbook(self.excelName, formatting_info=True)
        except FileNotFoundError:
            workbook = xlwt.Workbook(encoding="utf-8")  # 创建对象
            workbook.add_sheet("sheet1")
            workbook.save(self.excelName)
        finally:
            newbook = xlrd.open_workbook(self.excelName, formatting_info=True)
            sheet = newbook.sheets()[0]
            exist_rows = sheet.nrows
            exist_cols = sheet.ncols
            wtbook = copy(newbook)
            wtsheet = wtbook.get_sheet(0)
            for i, tr in enumerate(trs):
                tdlist = tr.select("td")
                for j, td in enumerate(tdlist):
                    wtsheet.write(exist_rows, j, td.text)
                exist_rows += 1
            wtbook.save(self.excelName)
