# 訂單處理腳本

紀錄2022年時製作的腳本，用於處理更新訂單總表。

Table of contents:

- [訂單處理腳本](#訂單處理腳本)
  - [01 執行流程](#01-執行流程)
  - [02 其他](#02-其他)

## 01 執行流程

1. Import `action/web_get.py`，並執行登入訂單網站後臺，進行資料爬取。
   1. 爬取完成的資料會先暫存到`excel_place`
2. Import `action/copy_act.py`，將資料更新到公用資料夾中的Excel總表。

## 02 其他

- 訂單網址為內網後臺
