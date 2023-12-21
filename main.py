
import asyncio
from pyppeteer import launch
import time
from bs4 import BeautifulSoup
import asyncio
import nest_asyncio
import datetime
import sqlite3
import pandas as pd
nest_asyncio.apply()

DBPATH = 'D:/Python/test.db'

trlist =[]

LottoryList =[]

class LottoryNum:
    def __init__(self, ID,Date, Num1,Num2,Num3,Num4,Num5,TopPriceNum):
        self.ID = ID
        self.Date = Date
        self.Num1 = Num1
        self.Num2 = Num2
        self.Num3 = Num3
        self.Num4 = Num4
        self.Num5 = Num5
        self.TopPriceNum = TopPriceNum
    
    def sortnum(self):
        nums = [self.Num1, self.Num2, self.Num3, self.Num4, self.Num5]
        nums.sort()
        self.Num1, self.Num2, self.Num3, self.Num4, self.Num5 = nums

def getlottorylist():
    return LottoryList

def gettrlist():
    return trlist

#取得資料
async def get_page_content(url,year,month):
    browser = await launch(headless = True)
    page = await browser.newPage()
    await page.goto(url, {"waitUntil": "networkidle0"})
    await page.click('#D539Control_history1_radYM') # 點擊 id = D539Control_history1_radYM
    await asyncio.sleep(3) # 等待 2 秒鐘

    # 選擇要查詢的年份
    select_year = str(year)
    # 點擊下拉選單
    await page.waitForSelector('#D539Control_history1_dropYear')
    # 點擊指定的月份
    await page.select('#D539Control_history1_dropYear', str(select_year))

    # 選擇要查詢的月份，例如選擇2月份
    selected_month = str(month)
    # 點擊下拉選單
    await page.waitForSelector('#D539Control_history1_dropMonth')
    # 點擊指定的月份
    await page.select('#D539Control_history1_dropMonth', str(selected_month))
    
    await asyncio.sleep(3) # 等待 2 秒鐘
    await page.click('#D539Control_history1_btnSubmit') # 點擊 id = D539Control_history1_btnSubmit
    # 等待網頁更新
    #await page.waitForNavigation()
    await asyncio.sleep(3)
    # 等待一段時間，讓網頁內容更新完畢
    html_content = await page.content()
    await browser.close()
    return html_content

async def getdata(year,month):
    url = "https://www.taiwanlottery.com.tw/lotto/dailycash/history.aspx"
    html_content = await get_page_content(url,year,month)
    soup = BeautifulSoup(html_content, "html.parser")
    table = soup.find('table', {'id': 'D539Control_history1_dlQuery'})
    tbody = table.find('tbody')
    print('End')
    return tbody

async def settrlist(page_tex):
    
    for tr in page_tex.children:
        if isinstance(tr, str) or tr.name != 'tr':
            continue
        else:
            trlist.append(tr)


def test():
    with sqlite3.connect(DBPATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM L539")

        results = cursor.fetchall()

        for data in results:
            print(data)

#修改抓下來的日期
def changData(Data):
    print("{yy}-{mm}-{dd}".format(yy=1911+int(Data[0]), mm=Data[1],dd=Data[2]))
    return("{yy}-{mm}-{dd}".format(yy=1911+int(Data[0]), mm=Data[1],dd=Data[2]))

async def setdata():
    datalist = gettrlist()
    for i in range(0,len(datalist)):
        tds = datalist[i].find_all('td')
        stringdata = tds[5].span.text.split('/')
        setData = changData(stringdata)
        tempdata = LottoryNum(tds[4].text.strip(),setData,tds[7].text,tds[8].text,tds[9].text,tds[10].text,tds[11].text,tds[29].text)
        tempdata.sortnum()
        LottoryList.append(tempdata)

async def setdatatosql():
    with sqlite3.connect(DBPATH) as conn:
        cursor = conn.cursor()
        # 確認資料庫是否被鎖定
        while True:
            try:
                cursor.execute("SELECT ID FROM L539")
                break
            except sqlite3.OperationalError:
                print("資料庫被鎖定，等待一秒鐘後重試...")
                time.sleep(1)
        # 取得查詢結果
        results = cursor.fetchall()
    
        existing_ids = set()
        
        for row in results:
            existing_ids.add(row[0])
        
        # 將多筆資料加入資料庫
        for d in LottoryList:
            stripped_data = int(d.ID.strip())
            if stripped_data not in existing_ids:
                print(stripped_data)
                cursor.execute('''INSERT INTO L539 (ID, Date, Num1, Num2, Num3, Num4, Num5, TopPriceNum) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', 
                                (int(d.ID.strip()), d.Date, d.Num1, d.Num2, d.Num3, d.Num4, d.Num5, d.TopPriceNum))
        
        # 提交資料庫更改並關閉連線
        conn.commit()
    
    print('SetDataToSqlEnd')

async def setdatatoExcel():
    data = {
    'ID': [item.ID for item in LottoryList],
    'Date': [item.Date for item in LottoryList],
    'Num1': [item.Num1 for item in LottoryList],
    'Num2': [item.Num2 for item in LottoryList],
    'Num3': [item.Num3 for item in LottoryList],
    'Num4': [item.Num4 for item in LottoryList],
    'Num5': [item.Num5 for item in LottoryList],
    'TopPriceNum': [item.TopPriceNum for item in LottoryList]
}
    df = pd.DataFrame(data)

    # 寫入 Excel 檔案
    excel_file_path = 'output_data.xlsx'
    df.to_excel(excel_file_path, index=False, engine='openpyxl')

    print(f'資料已匯出至 {excel_file_path}')

async def main():
    print('開始執行')

    for year in range(112, 113): # 範圍是100~112年
        for month in range(1, 12): # 範圍是1~12月
            print(f'目前正在抓取{year}年{month}月的資料')
            page_tex = await getdata(year,month)
            await settrlist(page_tex)
            await setdata()
            await setdatatoExcel()
            time.sleep(10)
    



if __name__ == '__main__':
    asyncio.run(main())
    #test()