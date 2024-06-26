# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : ask_helper.py
# @Software : PyCharm


import time

import requests
import re
import bs4
import xlwt


def recircle_get_dat():
    all = []
    for i in range(176):
        onehtml = get_doc_data(str(i + 1))
        onetab = parse_data(onehtml)
        print(i, onetab)
        for item in onetab:
            all.append(item)
        time.sleep(1)
    return all




def parse_data(html):
    soup = bs4.BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="myTable04")
    result = []
    for tr in table.find_all("tr"):
        links = tr.find_all(href=re.compile('summary'),  string=re.compile('[^\\d]\\w'))
        for link in links:
            code = re.compile('summary/(\\d+)').findall(str(link))
            if not code:
                continue
            text = link.getText()
            if re.search('^(?!.*退)(?!.*\*ST).*', text):
                item = {"code": code[0], "shortname": text}
                result.append(item)
    return result


def get_doc_data(page):
    url = "https://s.askci.com/stock/a-0-0/" + page
    try:
        response = requests.get(url, timeout=30000)
        html = response.text
        return html
    except Exception as e:
        print("http_err:", e)


def save_excel(list):
    workbook = xlwt.Workbook(encoding="utf-8")  # 创建对象
    worksheet = workbook.add_sheet("sheet1")
    for i, dict in enumerate(list):
        print(i, dict)
        worksheet.write(i, 0, dict['code'])  # 写入内容 第一个参数是行，第二个是列，第三个是内容
        worksheet.write(i, 1, dict['shortname'])
    workbook.save("ACode.xls")


if __name__ == '__main__':
    list_dict = recircle_get_dat()
    save_excel(list_dict)
    # save_excel(list)
