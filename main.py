# This is a sample Python script.
import all_code
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import ask_helper
#import pandas_financial as pdfn
#import stock_mysql_handler


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    '''
    测试功能
    '''
    ##list_dict = all_code.recircle_get_dat()
    ##all_code.save_excel(list_dict)

    ##ask_helper.get_financial_excel()

    '''
    获取现金流量表，利润表，资产负债表
    '''
    ask_helper.get_three_excel()

    '''
    #从3个表中获取数据后分析 并生产折线图
    '''
    ##pdfn.get_normal_data()

    '''
    从3个表中获取数据整理并存库
    '''
    ##stock_mysql_handler.prepare_cashflow_data_all()

