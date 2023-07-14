import pandas as pd
import pynput.mouse as ms
import pynput.keyboard as kb
import pyautogui as pag
import pyperclip as pc
import time

myMouse = ms.Controller()
myKeyboard = kb.Controller()

#file-->utf-8 
df = pd.read_csv("output.csv", encoding="utf-8")
df.head(30)

#use purchase in
cur_dic= {"USD": 0, "CNY": 0, "HKD": 0, "JPY": 0}
for i in cur_dic:
    cur_dic[i]= df[df["currency_code"]==i].purchase_in.item()
print(cur_dic)

#transfer exchange rate
#USD
U = 1 / cur_dic['USD']
U2 = round(U, 9)

#CNY
C = 1 / cur_dic['CNY']
C2 = round(C, 9)

#HKD
H = 1 / cur_dic['HKD']
H2 = round(H, 9)

#JPY
J = 1 / cur_dic['JPY']
J2 = round(J, 9)

print('USD', U2)
print('CNY', C2)
print('HKD', H2)
print('JPY', J2)

def Autoupdate_er():
    #search the program
    time.sleep(3)
    myMouse.position = (91, 1049)
    myMouse.click(ms.Button.left, 1)
    time.sleep(4)

    #search and open chrome
    # programToFind = input('Type the program name:')
    programToFind = 'chrome'
    myKeyboard.type(programToFind)
    time.sleep(4)
    myKeyboard.tap(kb.Key.enter)
    time.sleep(8)

    #click user
    myMouse.position = (944, 645)
    myMouse.click(ms.Button.left, 1)
    time.sleep(5)

    #click system
    myMouse.position = (186, 108)
    myMouse.click(ms.Button.left, 1)
    time.sleep(13)

    #click 設定
    myMouse.position = (1818, 156)
    myMouse.click(ms.Button.left, 1)
    time.sleep(8)

    #click Company Details
    myMouse.position = (257, 422)
    myMouse.click(ms.Button.left, 1)
    time.sleep(20)

    #click Currencies
    myMouse.position = (734, 241)
    myMouse.click(ms.Button.left, 1)
    time.sleep(10)

    #open USD and update
    myMouse.position = (478, 540)
    myMouse.click(ms.Button.left, 1)
    time.sleep(15)
    myMouse.position = (980, 394)
    myMouse.click(ms.Button.left, 2)
    time.sleep(3)

    a = U2
    pc.copy(a)
    pag.hotkey('ctrl', 'v')

    myMouse.position = (1180, 487)
    myMouse.click(ms.Button.left, 1)
    time.sleep(18)

    #open CNY and update
    myMouse.position = (515, 613)
    myMouse.click(ms.Button.left, 1)
    time.sleep(8)
    myMouse.position = (980, 394)
    myMouse.click(ms.Button.left, 2)
    time.sleep(3)

    b = C2
    pc.copy(b)
    pag.hotkey('ctrl', 'v')

    myMouse.position = (1180, 487)
    myMouse.click(ms.Button.left, 1)
    time.sleep(10)

    #open HKD and update
    myMouse.position = (512, 668)
    myMouse.click(ms.Button.left, 1)
    time.sleep(8)
    myMouse.position = (980, 394)
    myMouse.click(ms.Button.left, 2)
    time.sleep(3)

    c = H2
    pc.copy(c)
    pag.hotkey('ctrl', 'v')

    myMouse.position = (1180, 487)
    myMouse.click(ms.Button.left, 1)
    time.sleep(10)

    #open JPY and update
    myMouse.position = (462, 725)
    myMouse.click(ms.Button.left, 1)
    time.sleep(8)
    myMouse.position = (980, 394)
    myMouse.click(ms.Button.left, 2)
    time.sleep(3)

    d = J2
    pc.copy(d)
    pag.hotkey('ctrl', 'v')
    myMouse.position = (1180, 487)
    myMouse.click(ms.Button.left, 1)
    time.sleep(15)

    #refresh
    myMouse.position = (105, 70)
    myMouse.click(ms.Button.left, 1)
    time.sleep(15)

    #close 
    myMouse.position = (1898, 21)
    myMouse.click(ms.Button.left, 1)
    time.sleep(3)

Autoupdate_er()
