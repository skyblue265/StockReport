import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook


def CreateExcelFile(path, fileName):
    wb = Workbook()
    ws = wb.active
    ws.title = fileName

    wb.save(filename=f"{path}\\{fileName}.xlsx")
    return f"{path}\\{fileName}.xlsx"


def WriteToExcel(excelPath, worksheet, x, y, value):
    workbook = openpyxl.load_workbook(excelPath)
    worksheet = workbook[worksheet]
    worksheet[f"{x}{y}"] = value
    workbook.save(excelPath)


def CheckStockExist(stockNo):
    cookies = {
        'jcsession': 'jHttpSession^@7ea29a39',
        '_ga': 'GA1.3.1202029403.1602070641',
        '_gid': 'GA1.3.1184316053.1605093422',
        'newmops2': 'selfObj^%^3DtagCon1^%^7Cco_id^%^3D2448^%^7Cyear^%^3D^%^7Cseason^%^3D^%^7C',
        '_gat': '1',
    }

    headers = {
        'Connection': 'keep-alive',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': '*/*',
        'Origin': 'https://mops.twse.com.tw',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://mops.twse.com.tw/mops/web/t05st32',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    }

    data = {
        'encodeURIComponent': '1',
        'step': '1',
        'firstin': '1',
        'off': '1',
        'keyword4': '',
        'code1': '',
        'TYPEK2': '',
        'checkbtn': '',
        'queryName': 'co_id',
        'inpuType': 'co_id',
        'TYPEK': 'all',
        'isnew': 'true',
        'co_id': stockNo,
        'year': '',
        'season': ''
    }

    response = requests.post('https://mops.twse.com.tw/mops/web/ajax_t05st32', headers=headers, cookies=cookies,
                             data=data)
    response.encoding = 'big-5'
    if '不繼續公開發行' not in response.text:
        return True
    else:
        return False

def DownLoadData(stockNo, excelPath, workSheetName):
# excelPath = 'C:\\Users\\user\\Desktop\\data.xlsx'
# workSheetName = '1'

    cookies = {
        'jcsession': 'jHttpSession^@7ea29a39',
        '_ga': 'GA1.3.1202029403.1602070641',
        '_gid': 'GA1.3.1184316053.1605093422',
        'newmops2': 'selfObj^%^3DtagCon1^%^7Cco_id^%^3D2448^%^7Cyear^%^3D^%^7Cseason^%^3D^%^7C',
        '_gat': '1',
    }

    headers = {
        'Connection': 'keep-alive',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': '*/*',
        'Origin': 'https://mops.twse.com.tw',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://mops.twse.com.tw/mops/web/t05st32',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    }

    data = {
        'encodeURIComponent': '1',
        'step': '1',
        'firstin': '1',
        'off': '1',
        'keyword4': '',
        'code1': '',
        'TYPEK2': '',
        'checkbtn': '',
        'queryName': 'co_id',
        'inpuType': 'co_id',
        'TYPEK': 'all',
        'isnew': 'true',
        'co_id': stockNo,
        'year': '',
        'season': ''
    }

    response = requests.post('https://mops.twse.com.tw/mops/web/ajax_t05st32', headers=headers, cookies=cookies, data=data)
    response.encoding = 'big-5'
    # print(response.text)
    soup = BeautifulSoup(response.text, 'html.parser')
    # allTD = soup.find_all('center')

    # for item in allTD:
    #     print(item)

    alltblHead = soup.find_all("td", class_="tblHead")
    for item in alltblHead:
        # print(item.text.strip())
        WriteToExcel(excelPath, workSheetName, 'D', 2, item.text.strip())

    allTD = soup.find_all('b')
    for item in range(len(allTD)):
        if item == 2:
            WriteToExcel(excelPath, workSheetName, 'A', 1, allTD[item].text.strip())
        if item == 3:
            WriteToExcel(excelPath, workSheetName, 'A', 3, allTD[item].text.strip())
        if item == 4:
            WriteToExcel(excelPath, workSheetName, 'B', 3, allTD[item].text.strip())
        if item == 5:
            WriteToExcel(excelPath, workSheetName, 'D', 3, allTD[item].text.strip())
        if item == 6:
            WriteToExcel(excelPath, workSheetName, 'B', 4, allTD[item].text.strip())
        if item == 7:
            WriteToExcel(excelPath, workSheetName, 'C', 4, allTD[item].text.strip())
        if item == 8:
            WriteToExcel(excelPath, workSheetName, 'D', 4, allTD[item].text.strip())
        if item == 9:
            WriteToExcel(excelPath, workSheetName, 'E', 4, allTD[item].text.strip())

    allTD = soup.find_all('td')
    # for item in allTD:
    #     print(item)

    for item in range(len(allTD)):
        # if allTD[item].text.strip() == '營業外費用及損失':
        #      print(f"a{allTD[item].text}b")


        if allTD[item].text.strip() == '銷貨收入總額':
            WriteToExcel(excelPath, workSheetName, 'A', 5, '          銷貨收入總額')
            WriteToExcel(excelPath, workSheetName, 'B', 5, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 5, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 5, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 5, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '銷貨退回':
            WriteToExcel(excelPath, workSheetName, 'A', 6, '          銷貨退回')
            WriteToExcel(excelPath, workSheetName, 'B', 6, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 6, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 6, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 6, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '銷貨折讓':
            WriteToExcel(excelPath, workSheetName, 'A', 7, '          銷貨折讓')
            WriteToExcel(excelPath, workSheetName, 'B', 7, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 7, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 7, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 7, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '銷貨收入淨額':
            WriteToExcel(excelPath, workSheetName, 'A', 8, '          銷貨收入淨額')
            WriteToExcel(excelPath, workSheetName, 'B', 8, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 8, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 8, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 8, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '營業收入合計':
            WriteToExcel(excelPath, workSheetName, 'A', 9, '          營業收入合計')
            WriteToExcel(excelPath, workSheetName, 'B', 9, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 9, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 9, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 9, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '銷貨成本':
            WriteToExcel(excelPath, workSheetName, 'A', 10, '          銷貨成本')
            WriteToExcel(excelPath, workSheetName, 'B', 10, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 10, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 10, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 10, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '營業成本合計':
            WriteToExcel(excelPath, workSheetName, 'A', 11, '          營業成本合計')
            WriteToExcel(excelPath, workSheetName, 'B', 11, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 11, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 11, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 11, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '營業毛利(毛損)':
            WriteToExcel(excelPath, workSheetName, 'A', 12, '          營業毛利(毛損)')
            WriteToExcel(excelPath, workSheetName, 'B', 12, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 12, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 12, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 12, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '聯屬公司間未實現利益':
            WriteToExcel(excelPath, workSheetName, 'A', 13, '          聯屬公司間未實現利益')
            WriteToExcel(excelPath, workSheetName, 'B', 13, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 13, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 13, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 13, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '推銷費用':
            WriteToExcel(excelPath, workSheetName, 'A', 14, '          推銷費用')
            WriteToExcel(excelPath, workSheetName, 'B', 14, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 14, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 14, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 14, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '管理及總務費用':
            WriteToExcel(excelPath, workSheetName, 'A', 15, '          管理及總務費用')
            WriteToExcel(excelPath, workSheetName, 'B', 15, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 15, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 15, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 15, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '研究發展費用':
            WriteToExcel(excelPath, workSheetName, 'A', 16, '          研究發展費用')
            WriteToExcel(excelPath, workSheetName, 'B', 16, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 16, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 16, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 16, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '營業費用合計':
            WriteToExcel(excelPath, workSheetName, 'A', 17, '          營業費用合計')
            WriteToExcel(excelPath, workSheetName, 'B', 17, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 17, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 17, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 17, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '營業淨利(淨損)':
            WriteToExcel(excelPath, workSheetName, 'A', 18, '          營業淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 18, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 18, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 18, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 18, allTD[item + 4].text.strip())

        if allTD[item].text == '        營業外收入及利益':
            WriteToExcel(excelPath, workSheetName, 'A', 19, '        營業外收入及利益')
            WriteToExcel(excelPath, workSheetName, 'B', 19, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 19, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 19, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 19, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '利息收入':
            WriteToExcel(excelPath, workSheetName, 'A', 20, '          利息收入')
            WriteToExcel(excelPath, workSheetName, 'B', 20, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 20, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 20, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 20, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '投資收益':
            WriteToExcel(excelPath, workSheetName, 'A', 21, '          投資收益')
            WriteToExcel(excelPath, workSheetName, 'B', 21, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 21, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 21, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 21, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '股利收入':
            WriteToExcel(excelPath, workSheetName, 'A', 22, '          股利收入')
            WriteToExcel(excelPath, workSheetName, 'B', 22, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 22, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 22, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 22, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '處分投資利益':
            WriteToExcel(excelPath, workSheetName, 'A', 23, '          處分投資利益')
            WriteToExcel(excelPath, workSheetName, 'B', 23, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 23, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 23, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 23, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '兌換利益':
            WriteToExcel(excelPath, workSheetName, 'A', 24, '          兌換利益')
            WriteToExcel(excelPath, workSheetName, 'B', 24, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 24, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 24, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 24, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '租金收入':
            WriteToExcel(excelPath, workSheetName, 'A', 25, '          租金收入')
            WriteToExcel(excelPath, workSheetName, 'B', 25, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 25, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 25, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 25, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '金融資產評價利益':
            WriteToExcel(excelPath, workSheetName, 'A', 26, '          金融資產評價利益')
            WriteToExcel(excelPath, workSheetName, 'B', 26, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 26, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 26, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 26, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '金融資產評價利益':
            WriteToExcel(excelPath, workSheetName, 'A', 27, '          金融資產評價利益')
            WriteToExcel(excelPath, workSheetName, 'B', 27, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 27, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 27, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 27, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '什項收入':
            WriteToExcel(excelPath, workSheetName, 'A', 28, '          什項收入')
            WriteToExcel(excelPath, workSheetName, 'B', 28, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 28, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 28, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 28, allTD[item + 4].text.strip())

        if allTD[item].text == '          營業外收入及利益':
            WriteToExcel(excelPath, workSheetName, 'A', 29, '          營業外收入及利益')
            WriteToExcel(excelPath, workSheetName, 'B', 29, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 29, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 29, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 29, allTD[item + 4].text.strip())

        if allTD[item].text == '        營業外費用及損失':
            WriteToExcel(excelPath, workSheetName, 'A', 30, '        營業外費用及損失')
            WriteToExcel(excelPath, workSheetName, 'B', 30, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 30, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 30, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 30, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '利息費用':
            WriteToExcel(excelPath, workSheetName, 'A', 31, '          利息費用')
            WriteToExcel(excelPath, workSheetName, 'B', 31, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 31, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 31, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 31, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '採權益法認列之投資損失':
            WriteToExcel(excelPath, workSheetName, 'A', 32, '          採權益法認列之投資損失')
            WriteToExcel(excelPath, workSheetName, 'B', 32, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 32, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 32, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 32, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '投資損失':
            WriteToExcel(excelPath, workSheetName, 'A', 33, '          投資損失')
            WriteToExcel(excelPath, workSheetName, 'B', 33, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 33, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 33, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 33, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '處分固定資產損失':
            WriteToExcel(excelPath, workSheetName, 'A', 34, '          處分固定資產損失')
            WriteToExcel(excelPath, workSheetName, 'B', 34, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 34, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 34, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 34, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '處分投資損失':
            WriteToExcel(excelPath, workSheetName, 'A', 35, '          處分投資損失')
            WriteToExcel(excelPath, workSheetName, 'B', 35, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 35, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 35, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 35, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '兌換損失':
            WriteToExcel(excelPath, workSheetName, 'A', 36, '          兌換損失')
            WriteToExcel(excelPath, workSheetName, 'B', 36, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 36, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 36, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 36, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '減損損失':
            WriteToExcel(excelPath, workSheetName, 'A', 37, '          減損損失')
            WriteToExcel(excelPath, workSheetName, 'B', 37, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 37, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 37, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 37, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '金融資產評價損失':
            WriteToExcel(excelPath, workSheetName, 'A', 38, '          金融資產評價損失')
            WriteToExcel(excelPath, workSheetName, 'B', 38, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 38, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 38, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 38, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '金融負債評價損失':
            WriteToExcel(excelPath, workSheetName, 'A', 39, '          金融負債評價損失')
            WriteToExcel(excelPath, workSheetName, 'B', 39, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 39, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 39, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 39, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '什項支出':
            WriteToExcel(excelPath, workSheetName, 'A', 40, '          什項支出')
            WriteToExcel(excelPath, workSheetName, 'B', 40, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 40, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 40, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 40, allTD[item + 4].text.strip())

        if allTD[item].text == '          營業外費用及損失':
            WriteToExcel(excelPath, workSheetName, 'A', 41, '          營業外費用及損失')
            WriteToExcel(excelPath, workSheetName, 'B', 41, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 41, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 41, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 41, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '繼續營業單位稅前淨利(淨損)':
            print(item)
            WriteToExcel(excelPath, workSheetName, 'A', 42, '          繼續營業單位稅前淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 42, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 42, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 42, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 42, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '所得稅費用(利益)':
            WriteToExcel(excelPath, workSheetName, 'A', 43, '          所得稅費用(利益)')
            WriteToExcel(excelPath, workSheetName, 'B', 43, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 43, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 43, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 43, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '繼續營業單位淨利(淨損)' and item == 201:
            WriteToExcel(excelPath, workSheetName, 'A', 44, '          繼續營業單位淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 44, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 44, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 44, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 44, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '本期淨利(淨損)':
            WriteToExcel(excelPath, workSheetName, 'A', 45, '          本期淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 45, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 45, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 45, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 45, allTD[item + 4].text.strip())

        if allTD[item].text == '        基本每股盈餘':
            WriteToExcel(excelPath, workSheetName, 'A', 46, '        基本每股盈餘')
            WriteToExcel(excelPath, workSheetName, 'B', 46, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 46, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 46, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 46, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '繼續營業單位淨利(淨損)' and item == 216:
            print(item)
            WriteToExcel(excelPath, workSheetName, 'A', 47, '          繼續營業單位淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 47, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 47, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 47, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 47, allTD[item + 4].text.strip())

        if allTD[item].text == '          基本每股盈餘':
            WriteToExcel(excelPath, workSheetName, 'A', 48, '          基本每股盈餘')
            WriteToExcel(excelPath, workSheetName, 'B', 48, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 48, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 48, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 48, allTD[item + 4].text.strip())

        if allTD[item].text == '        稀釋每股盈餘':
            WriteToExcel(excelPath, workSheetName, 'A', 49, '        稀釋每股盈餘')
            WriteToExcel(excelPath, workSheetName, 'B', 49, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 49, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 49, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 49, allTD[item + 4].text.strip())

        if allTD[item].text.strip() == '繼續營業單位淨利(淨損)' and item == 231:
            print(item)
            WriteToExcel(excelPath, workSheetName, 'A', 50, '          繼續營業單位淨利(淨損)')
            WriteToExcel(excelPath, workSheetName, 'B', 50, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 50, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 50, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 50, allTD[item + 4].text.strip())

        if allTD[item].text == '          稀釋每股盈餘':
            WriteToExcel(excelPath, workSheetName, 'A', 51, '          稀釋每股盈餘')
            WriteToExcel(excelPath, workSheetName, 'B', 51, allTD[item + 1].text.strip())
            WriteToExcel(excelPath, workSheetName, 'C', 51, allTD[item + 2].text.strip())
            WriteToExcel(excelPath, workSheetName, 'D', 51, allTD[item + 3].text.strip())
            WriteToExcel(excelPath, workSheetName, 'E', 51, allTD[item + 4].text.strip())

for i in range(1000,9999):
    #檢查股票是否存在

    StockIsExist = CheckStockExist(i)

    if StockIsExist == True:
    #產生EXCEL檔案
        filePath = CreateExcelFile("C:\\Users\\user\\Desktop", i)
    # 對excel檔案寫入資料
        DownloadData(i, filePath, i)
        Time.sleep(10)
