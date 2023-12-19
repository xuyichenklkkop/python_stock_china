import time

import requests
import re
import xlrd
import bs4
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
        if re.match("\\*ST", one_sheet.row_values(i)[1]):
            continue
        result.append({one_sheet.row_values(i)[0]: one_sheet.row_values(i)[1]})
    return result


def get_stock_control_man():
    codes = read_excel("ACode.xls")
    list =[]
    for code in codes:
        for key, name in code.items():
            onehtml = get_stock_data(key)
            item= parse_data(onehtml, key, name)
            list.append(item)
            time.sleep(0.5)

    save_excel(list)


def get_stock_data(code):
    url = "https://s.askci.com/stock/summary/" + code
    try:
        response = requests.get(url, timeout=30000)
        response.encoding = 'utf-8'
        html = response.text
        return html
    except Exception as e:
        print("http_err:", e)


def generate_patterns(words):
    # 将多个单词生成对应的正则表达式
    patterns = tuple(re.compile(r'\b{}\b'.format(word)) for word in words)
    return patterns


def parse_data(html, key, name):
    bs = bs4.BeautifulSoup(html, "html.parser")
    trs = bs.select("div[class='right_f_c_table mg_tone'] > table > tr")
    print(key, name)
    item ={'code':key,'name':name}
    for index, row in enumerate(trs):
        shiji = row.findChild('td', string=re.compile('实际控制人'))
        if shiji:
            td = shiji.find_next_sibling()
            item['shiji'] = td.text
            print(td.text)
        zui = row.findChild('td', string=re.compile('最终控制人'))
        if zui:
            td = zui.find_next_sibling()
            item['zuizhong'] = td.text
            print(td.text)
    return item

def save_excel(list):
    workbook = xlwt.Workbook(encoding="utf-8")  # 创建对象
    worksheet = workbook.add_sheet("sheet1")
    for i,dict in enumerate(list):
        print(i,dict)
        worksheet.write(i, 0, dict['code'])  # 写入内容 第一个参数是行，第二个是列，第三个是内容
        worksheet.write(i, 1, dict['name'])
        worksheet.write(i, 2, dict['shiji'])
        worksheet.write(i, 3, dict['zuizhong'])
    workbook.save("stock_owner.xls")




if __name__ == '__main__':
    # onehtml = get_stock_data('000068')
    # parse_data(onehtml,'000068','赛格')
    get_stock_control_man()
