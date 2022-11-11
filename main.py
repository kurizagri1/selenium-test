import os
from time import sleep

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
url = 'neural-university.ru'
driver.get("https://www.similarweb.com/ru/website/" + url + "/#overview")
sleep(20)  # bypass site protection
elems = driver.find_elements(By.CLASS_NAME, "wa-competitors__list-item-title")
url_list = []
for elem in elems:
    url_list.append(elem.text)

for url in url_list:
    driver = webdriver.Chrome()
    driver.get("https://www.similarweb.com/ru/website/" + url + "/#overview")
    sleep(20)
    elem = driver.find_element(By.CLASS_NAME, "engagement-list__item-value")
    wb = load_workbook('data.xlsx')
    ws = wb.active
    ws.append([elem.text])
    wb.save("data.xlsx")
    driver.close()

os.system('"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" data.xlsx')
