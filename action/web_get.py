"""
import selenium, webdriver(要更新), numpy(非必要), pandas
此檔案可以連到Dolibarr 5.0.1, 並登入網站, 搜尋指定訂單並複製內容, 產生Excel檔案
檔名用訂單號碼命名, sheet1用客戶代號命名
"""
import sys
from operator import index
from textwrap import fill
from webbrowser import get
from selenium import webdriver # import webdriver
from selenium.webdriver.common.by import By # import find.element.by功能
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import time # import time.sleep(second)
import pandas as pd # import pandas處理資料
import numpy as np
import os
import openpyxl as opxl # import openpyxl 處理excel

# 忽略並隱藏 0x1F error
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

print('** 請勿輸入草稿訂單網址')
order_url = input('請輸入訂單網址')
driver = webdriver.Chrome('./chromedriver/chromedriver.exe',options=options)
# driver.maximize_window()
driver.minimize_window()
print('please wait a minute...')
try: 
    driver.get(order_url)
except:
    print('請輸入正確網址')
    sys.exit()

username = "請輸入"
password = "請輸入"

# 輸入帳密
username_textfield = driver.find_element(By.NAME, 'username')
password_textfield = driver.find_element(By.NAME, "password")
login_button = driver.find_element(By.CLASS_NAME, 'button')

# 登入
username_textfield.send_keys(username)
password_textfield.send_keys(password)
login_button.click()
time.sleep(1)

# find PO#, 客戶名稱/代號
po_find = driver.find_element(By.XPATH, '//*[@id="id-right"]/div/div[2]/div[1]/div/div[4]').text
po_number = po_find[po_find.find('CF'):po_find.find('供應商編號')-1]
supplier_number = po_find[po_find.find('SU'):po_find.find('客戶/供應商')-1]
supplier_name = po_find[po_find.find('客戶/供應商')+9:]
# print PO號碼
print('訂單編號:',po_number)
print('供應商編號:',supplier_number)
print('供應商名稱:',supplier_name)

# 抓取table
element = driver.find_element(By.XPATH, '//*[@id="tablelines"]/tbody')
# 抓取td
td_content = element.find_elements(By.TAG_NAME, 'td')
# lst list
lst = []
# 新增td文字 in lst
for td in td_content:
    lst.append(td.text)
    # remove空白元素
    for i in lst:
        if '' in lst:
            lst.remove('')

# 表格欄數抓取&設定
col = len(element.find_elements(By.XPATH, '//*[@id="tablelines"]/tbody/tr[2]/td')) - 1
# print(col)
lst = [lst[i:i + col] for i in range(0, len(lst), col)]

# find 訂單時間
po_time = driver.find_element(By.XPATH, '//*[@id="builddoc_form"]/table[2]/tbody/tr[2]/td[3]').text

# find 訂單交期
leadtime = driver.find_element(By.XPATH, '//*[@id="id-right"]/div/div[2]/div[3]/div[1]/table/tbody/tr[4]/td[2]').text

# find 備註
remark_link = driver.find_element(By.ID, 'note')
remark_link.click()
remark = driver.find_element(By.XPATH, '//*[@id="id-right"]/div/div[2]/div[3]/div[2]/div[1]/div[2]').text

# pandas設定list to dataframe
web_table = pd.DataFrame(lst)
web_table.columns = ['品名','營業稅','單價','數量','折扣','銷售金額']
drop_table = web_table.drop([0],axis=0)
drop_table.to_excel('./excel_place/%s.xlsx'%po_number,sheet_name=po_number)

# 離開webdriver
driver.quit()
print('資料擷取成功\n訂單: %s'%po_number)
