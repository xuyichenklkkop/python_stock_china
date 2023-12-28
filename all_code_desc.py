# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : ask_helper.py
# @Software : PyCharm


import time

import requests
import re
import bs4
import xlrd
import xlwt

MaxPage = 1  # the max page we don't know
Code_Execl_Name = "ACode.xls"

def circle_get_data():
    for i in range(MaxPage):
        page_items = []
        page_html = get_doc_data(str(i + 1))
        tabs = parse_data(page_html)
        print(i, tabs)
        for item in tabs:
            page_items.append(item)
        save_excel(page_items)


def parse_data(html):
    bs = bs4.BeautifulSoup(html, "html.parser")
    table = bs.select("table[id='myTable04']")
    heads = bs.select('thead > tr th')
    result = []
    for tb in table:
        trs = bs.select('tbody > tr')
        for tr in trs:
            tds = tr.find_all("td")
            for idx, td in enumerate(tds):
                print(td)

        # for child in tb.children:
        #     links = bs4.BeautifulSoup(str(child), 'html.parser').find_all(href=re.compile('summary'),
        #                                                                   string=re.compile('[^HK][^\\d]\\w'))
        #     for link in links:
        #         code = re.compile('summary/(\\d+)').findall(str(link))
        #         text = link.getText()
        #         if code:
        #             item = {"code": code[0], "shortname": text}
        #             result.append(item)
    return result


def get_doc_data(page):
    url = f"https://s.askci.com/stock/a/0-0?reportTime=2023-09-30&pageNum={page}#QueryCondition"
    try:
        response = requests.get(url, timeout=30000)
        html = response.text
        return html
    except Exception as e:
        print("http_err:", e)


def save_excel(page_lists):
    try:
        workbook_to_read = xlrd.open_workbook(Code_Execl_Name)
    except FileNotFoundError:
        workbook = xlwt.Workbook(encoding="utf-8")  # 创建对象
        one_sheet = workbook.add_sheet("sheet1")
        workbook.save(Code_Execl_Name)

    work_write_again = xlwt.Workbook(encoding="utf-8")
    sheet_again = work_write_again.add_sheet("sheet1")
    work_read_again = xlrd.open_workbook(Code_Execl_Name)
    existing_sheet = work_read_again.sheet_by_index(0)
    rows = existing_sheet.nrows
    for row in range(existing_sheet.nrows):
        for col in range(existing_sheet.ncols):
            sheet_again.write(row, col, existing_sheet.cell_value(row, col))  # exist data read and write again

    for i, item in enumerate(page_lists):  # write page new data
        sheet_again.write(rows + i, 0, item['code'])  # 写入内容 第一个参数是行，第二个是列，第三个是内容
        sheet_again.write(rows + i, 1, item['shortname'])
        sheet_again.write(rows + i, 2, item['shiji'])
        sheet_again.write(rows + i, 3, item['zuizhong'])
    work_write_again.save("stock_owner.xls")



if __name__ == '__main__':
    circle_get_data()
