from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib.request
import openpyxl as xl
import pyperclip as clp

wb = xl.load_workbook('order.xlsx')
ws = wb.active
lst = []
for r in ws.rows:
    if r[0].value is None:
        continue
    lst.append([])
    for c in r:
        lst[-1].append(c.value)
lst.pop(0)

driver = webdriver.Chrome()
driver.get("https://mail.google.com/mail/u/1/#inbox")
elem = driver.find_element_by_name("identifier")
elem.send_keys("jaehyun@gameberry.co.kr")
elem.send_keys(Keys.RETURN)
time.sleep(1)
elem1 = driver.find_element_by_name("password")
time.sleep(1)
elem1.send_keys("wldnro1357!")
elem1.send_keys(Keys.RETURN)
time.sleep(2)
elem3 = driver.find_element_by_css_selector(".T-I.T-I-KE.L3")
elem3.click()
time.sleep(3)

elem4 = driver.find_element_by_css_selector('.vO')
elem5 = driver.find_element_by_name('subjectbox')
elem6=  driver.find_element_by_id(':ox')
elem7= driver.find_element_by_css_selector('.T-I.J-J5-Ji.aoO.v7.T-I-atl.L3')

for i in lst:
    time.sleep(1)
    elem4.click()
    time.sleep(2)
    elem4.send_keys(i[-1])
    elem5.send_keys('[게임베리] {}님의 주문내역을 알려드립니다'.format(i[1]))
    elem6.send_keys('''
    안녕요 {}님
    {}-{}-{}에 주문하신 제품에 대해 알려줄게요
    제품명 {}
    금액: {}
    주문 감사요 ㅋㅋ
    '''.format(i[1], i[0].year, i[0].month, i[0].day, i[2], i[3]))
    time.sleep(2.5)
    elem7.click()
    time.sleep(2)
    elem3.click()
    time.sleep(2)



    

# for i in lst:
#     x, y = gui.locateCenterOnScreen('1.png')
#     gui.click(x,y)
#     time.sleep(1)
#     clp.copy(i[-1])
#     gui.hotkey('command','v')
#     time.sleep(1)
#     gui.hotkey('tab')
#     clp.copy('[게임베리] {}님의 주문내역을 알려드립니다'.format(i[1]))
#     gui.hotkey('command','v')
#     time.sleep(1)
#     gui.hotkey('tab')
#     clp.copy('''
#     안녕요 {}님
#     {}-{}-{}에 주문하신 제품에 대해 알려줄게요
#     제품명 {}
#     금액: {:,}
#     주문 감사요 ㅋㅋ
#     '''.format(i[1], i[0].year, i[0].month, i[0].day, i[2], i[3]))
#     gui.hotkey('command','v')
#     time.sleep(1) 
    
#     x,y = gui.locateCenterOnScreen('2.png', grayscale =True)
#     gui.click(x,y)
#     time.sleep(20)