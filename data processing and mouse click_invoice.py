#pip install pyautogui
#pip install pynput
import pandas as pd
import pynput.mouse as ms
import pynput.keyboard as kb
import pyautogui
import time
import datetime
import os

myMouse = ms.Controller()
myKeyboard = kb.Controller()

def books_integration(path_altra, path_yr, path_yd):
    altra = pd.read_csv(path_altra, encoding = 'utf-8')
    yr = pd.read_csv(path_yr, encoding = 'utf-8')
    yd = pd.read_csv(path_yd, encoding = 'utf-8')

    altra['帳別'] = 'ALTRA ELECTRIC COPORATION'
    no_altra = 1
    altra['代號'] = no_altra
    altra['出貨單號'] = altra['單據號碼'].apply(lambda x: '00{}-'.format(no_altra) + str(x))
    altra['幣別'] = altra['幣別'].replace('RMB', 'CNY')
    altra = altra[['帳別', '代號', '出貨單號', '單據日期', '單據', '單據號碼', '客戶簡稱', '客戶簡稱.1', '業務代碼', '業務名稱', '產品編號', '品名',
           '規格', '數量', '單位', '幣別', '匯率', '單價', '原幣金額', '台幣金額']]

    yd['帳別'] = 'x1'
    no_yd = 4
    yd['代號'] = no_yd
    yd['出貨單號'] = yd['單據號碼'].apply(lambda x: '00{}-'.format(no_yd) + str(x))
    yd['幣別'] = yd['幣別'].replace('RMB', 'CNY')
    yd = yd[['帳別', '代號', '出貨單號', '單據日期', '單據', '單據號碼', '客戶簡稱', '客戶簡稱.1', '業務代碼', '業務名稱', '產品編號', '品名',
           '規格', '數量', '單位', '幣別', '匯率', '單價', '原幣金額', '台幣金額']]

    yr['帳別'] = 'x2'
    no_yr = 6
    yr['代號'] = no_yr
    yr['出貨單號'] = yr['單據號碼'].apply(lambda x: '00{}-'.format(no_yr) + str(x))
    yr['幣別'] = yr['幣別'].replace('RMB', 'CNY')
    yr = yr[['帳別', '代號', '出貨單號', '單據日期', '單據', '單據號碼', '客戶簡稱', '客戶簡稱.1', '業務代碼', '業務名稱', '產品編號', '品名',
           '規格', '數量', '單位', '幣別', '匯率', '單價', '原幣金額', '台幣金額']]
    
    final_df = pd.concat([altra, yd], axis=0)
    final_df = pd.concat([final_df, yr], axis=0)
    return final_df

# file-->utf-8
path_altra = 'C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice\\001_Invoices.csv' 
path_yd = 'C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice\\004_Invoices.csv'
path_yr ='C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice\\006_Invoices.csv'
integrated = books_integration(path_altra, path_yr, path_yd)

integrated_出貨 = integrated.copy()
integrated_出貨明細 = integrated.copy()

integrated_出貨.rename(columns = {
    '單據日期': '出貨日期', '客戶簡稱.1': '客戶名稱',
    '業務名稱': '負責業務', '幣別': 'Currency', '匯率': '單據匯率'
}, inplace = True)

integrated_出貨明細.rename(columns={
    '出貨單號': '出貨(出貨單號)', '數量': '數量',
    '幣別': 'Currency', '單價': '單價', 
    '原幣金額': '小計', '客戶簡稱.1': '客戶名稱'
}, inplace=True)

integrated_出貨.to_csv('C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice\\出貨單_Upload.csv', encoding='cp950', index=False)
integrated_出貨明細.to_csv('C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice\\出貨明細_Upload.csv', encoding='cp950', index=False)

def AutoUploadInvoice():
    #search the program
    time.sleep(3)
    myMouse.position = (91, 1049)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #search and open chrome
    #program to find = input('type the program name:')
    programToFind = 'chrome'
    myKeyboard.type(programToFind)
    time.sleep(4)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(8)

    #click user
    myMouse.position = (961, 605)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

    #Import 出貨
    #choose module
    myMouse.position = (1200, 95)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

    programToFind = '出貨'
    myKeyboard.type(programToFind)
    time.sleep(4)
    myKeyboard.tap(kb.Key.enter)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(5)

    #click Import
    myMouse.position = (1751, 145)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Import 出貨
    myMouse.position = (1704, 196)
    myMouse.click(ms.Button.left, 1)    
    time.sleep(4)

    #click 瀏覽
    myMouse.position = (725, 426)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #input path
    myMouse.position = (654, 46)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)
    programToFind = 'C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice'
    myKeyboard.type(programToFind)
    myKeyboard.tap(kb.Key.enter)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(4)

    #input Import file name 
    myMouse.position = (275, 419)
    myMouse.click(ms.Button.left, 1)
    time.sleep(1)
    # myKeyboard.tap(kb.Key.shift)
    programToFind = '出貨單_Upload.csv'
    myKeyboard.type(programToFind)
    time.sleep(2)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(4)

    #click 下一步
    myMouse.position = (1341, 725)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #choose Both 
    myMouse.position = (365, 306)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click
    myMouse.position = (280, 353)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)   

    #click 出貨單號
    myMouse.position = (295, 457)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 下一步
    myMouse.position = (173, 473)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Apply auto mapping
    myMouse.position = (238, 981)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Apply
    myMouse.position = (1020, 231)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    myMouse.position = (417, 518)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    pyautogui.hotkey('ctrl', 'a')
    myKeyboard.press(kb.Key.backspace)
    programToFind = '客戶'
    myKeyboard.type(programToFind)

    myMouse.position = (432, 663)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 下一步
    myMouse.position = (1849, 987)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 繼續
    myMouse.position = (1015, 247)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    myMouse.scroll(0, -100)
    time.sleep(4)

    #click Finish
    myMouse.position = (280, 721)
    myMouse.click(ms.Button.left, 2)
    time.sleep(10)

    #click home and refresh
    myMouse.position = (260, 96)
    myMouse.click(ms.Button.left, 2)
    time.sleep(15)

    myMouse.position = (81, 51)
    myMouse.click(ms.Button.left, 2)
    myKeyboard.tap(kb.Key.f5)
    time.sleep(6)
AutoUploadInvoice()

def AutoUploadInvoiceDetail():
    #Import 出貨明細
    #choose module
    time.sleep(3)
    myMouse.position = (1200, 95)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

    programToFind = '出貨明細'
    myKeyboard.type(programToFind)
    time.sleep(4)
    myKeyboard.tap(kb.Key.enter)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(5)

    #click Import 出貨明細下拉選單
    myMouse.position = (1751, 145)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Import 出貨明細
    myMouse.position = (1704, 196)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 瀏覽
    myMouse.position = (725, 426)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #input path
    myMouse.position = (651, 51)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)
    programToFind = 'C:\\Users\\michelle.ho\\Desktop\\自動化\\Zoho CRM Invoice'
    myKeyboard.type(programToFind)
    myKeyboard.tap(kb.Key.enter)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(4)

    #input Import file name
    myMouse.position = (275, 419)
    myMouse.click(ms.Button.left, 1)
    time.sleep(1)
    # myKeyboard.tap(kb.Key.shift)
    programToFind = '出貨明細_Upload.csv'
    myKeyboard.type(programToFind)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(4)

    #click 下一步
    myMouse.position = (1342, 722)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    myMouse.position = (161, 441)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Apply auto mapping
    myMouse.position = (238, 981)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click Apply
    myMouse.position = (1020, 231)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 下一步
    myMouse.position = (1849, 987)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #click 繼續
    myMouse.position = (1015, 247)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    myMouse.scroll(0, -100)
    time.sleep(4)

    #click Finish
    myMouse.position = (280, 721)
    myMouse.click(ms.Button.left, 2)
    time.sleep(10)

    #click home and refresh
    myMouse.position = (260, 96)
    myMouse.click(ms.Button.left, 2)
    time.sleep(15)

    myMouse.position = (81, 51)
    myMouse.click(ms.Button.left, 2)
    myKeyboard.tap(kb.Key.f5)
    time.sleep(6)
AutoUploadInvoiceDetail()

def AutoConvertInvoice():
    #turn to標準出貨單
    #choose module
    time.sleep(5)
    myMouse.position = (1200, 95)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

    programToFind = '出貨'
    myKeyboard.type(programToFind)
    time.sleep(4)
    myKeyboard.tap(kb.Key.enter)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(5)

    #input Created Time 
    myMouse.position = (107, 284)
    myMouse.click(ms.Button.left, 1)
    programToFind = 'Created Time'
    myKeyboard.type(programToFind)
    time.sleep(2)
    myKeyboard.tap(kb.Key.enter)
    
    #click Created Time 
    myMouse.position = (38, 439)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)

    myMouse.position = (102, 460)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)

    myMouse.position = (116, 659)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)

    #click Apply Filter
    myMouse.position = (82, 956)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

    #click all Records
    myMouse.position = (388, 238)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)
 
    #回覆已轉入狀態
    myMouse.position = (680, 144)
    myMouse.click(ms.Button.left, 1)
    time.sleep(2)
 
    #click 轉入標準出貨單
    myMouse.position = (620, 239)
    myMouse.click(ms.Button.left, 1)
    time.sleep(15)
    
    #click home and refresh
    myMouse.position = (260, 96)
    myMouse.click(ms.Button.left, 2)
    time.sleep(15)

    myMouse.position = (81, 51)
    myMouse.click(ms.Button.left, 2)
    myKeyboard.tap(kb.Key.f5)
    time.sleep(6)
AutoConvertInvoice()

filePath1 = 'C:\\Users\\001_Invoices.csv'
filePath2 = 'C:\\Users\\004_Invoices.csv'
filePath3 = 'C:\\Users\\006_Invoices.csv'
filePath4 = 'C:\\Users\\出貨單_Upload.csv'
filePath5 = 'C:\\Users\\出貨明細_Upload.csv'

os.remove(filePath1)
os.remove(filePath2)
os.remove(filePath3)
os.remove(filePath4)
os.remove(filePath5)
