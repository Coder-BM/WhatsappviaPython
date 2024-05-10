import pyautogui as p
import pandas as pd
import keyboard as k
import webbrowser as web
import time as t
from urllib.parse import quote


def whatsapp(excel_file,text_file,x_cord=822,y_cord=1151):
    pf = pd.read_excel(excel_file,dtype={"Contact":str})
    name = pf["Name"].values
    contact = pf["Contact"].values
    zipper = zip(name,contact)

    files = text_file
    with open(files) as f:
        file = f.read()

    counter = 0

    for(a,b) in zipper:
        msg = file.format(a)
        web.open(f'https://web.whatsapp.com/send?phone={b}&text={quote(msg)}')
        t.sleep(10)
        p.click(x_cord,y_cord)
        t.sleep(5)
        k.press_and_release("enter")
        t.sleep(2)
        k.press_and_release("ctrl+w")
        t.sleep(2)
        counter += 1
        print(counter,"- Message Sent!!!!")
    print("Done") 

excel_path= r'C:\Users\Blaise\Desktop\Whats\Numbers.xlsx'
text_path = r'C:\Users\Blaise\Desktop\Whats\content.txt'

whatsapp(excel_path,text_path)   