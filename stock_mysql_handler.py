# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : stock_mysql_handler.py
# @Software : PyCharm

import mysql_helper
import pandas_financial


def prepare_cashflow_data_all():
    codes = pandas_financial.get_normal_code()
    for code in codes:
        for item in code.items():
            pdf = pandas_financial.pandas_finance()
            pdf.set_normal_code(item[0], item[1])
            pdf.set_read_file_name()

            cash_data = pdf.get_excel_cashflow_data()
            for item in cash_data:
                insert_cashflow_data_mysql_many(item)

            profit_data = pdf.get_excel_profit_all()
            print(profit_data)
            for item in profit_data:
                insert_profit_data_mysql_many(item)

            financial_data = pdf.get_excel_finical_all()
            for item in financial_data:
                insert_financial_data_mysql_many(item)


def insert_financial_data_mysql_many(tuplelist):
    with mysql_helper.MySqlHelper(logtime=True) as msh:
        sql = "insert into financial_report(stock_code,stock_name,report_date,cash,note_accounts_receivable,advances_to_suppliers,inventories," \
              "current_assets_all,fixed_assets_all,construction_in_progress" \
              ",intangible_assets,goodwill,longterm_deferred_expenses,noncurrent_assets_all," \
              "note_accounts_payable,current_liability_all,nocurrent_liability_all," \
              "share_capital,capital_reserves,surplus_reserves,undisturbed_profit) "
        val_l = len(tuplelist[0])
        s_count = val_l * "%s,"
        sql += " values (" + s_count[:-1] + ")"
        msh.cursor.executemany(sql, tuplelist)


def insert_cashflow_data_mysql_many(tuplelist):
    namelist = ['stock_code', 'stock_name', 'report_date', 'business_cash_in', 'business_cash_out',
                'business_cash_balance', 'sale_cash_in', 'purchase_cash_out']
    val_l = len(tuplelist[0])
    colname = ','.join(namelist[0:val_l])
    with mysql_helper.MySqlHelper(logtime=True) as msh:
        sql = "insert into cashflow_report ({})".format(colname)
        s_count = val_l * "%s,"
        sql += "values (" + s_count[:-1] + ")"
        msh.cursor.executemany(sql, tuplelist)


def insert_profit_data_mysql_many(tuplelist):
    with mysql_helper.MySqlHelper(logtime=True) as msh:
        sql = "insert into profit_report(stock_code,stock_name,report_date,income_all,cost_all,sale_cost,manage_cost,develop_cost,finance_cost," \
              "roi_profit,business_profit,noncurrent_assets_profit,noncurrent_assets_cost,net_profit,koufei_net_profit) "
        val_l = len(tuplelist[0])
        s_count = val_l * "%s,"
        sql += " values (" + s_count[:-1] + ")"
        msh.cursor.executemany(sql, tuplelist)
