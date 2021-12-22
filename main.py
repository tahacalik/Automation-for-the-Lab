from selenium import webdriver
import time
import xlrd
import numpy as np
import os

browser = "C:\edge.exe"
driver = webdriver.Edge(browser)
path = "C:\hsys.xlsx"

excel_workbook = xlrd.open_workbook(path)
excel_worksheet = excel_workbook.sheet_by_index(0)

for x in range (excel_worksheet.nrows):
    tc = (excel_worksheet.cell_value(x,2)) #3. sütun
    barkod = str((excel_worksheet.cell_value(x,3)))[:-2] #4. sütun
    isim = (excel_worksheet.cell_value(x,0)) #1. sütun
    driver.get('https://enabiz.gov.tr/PcrTestSonuc/Index')
    time.sleep(5)
    deneme = 0
    while deneme < 1:
                    try:
                       driver.execute_script(f"GetPcrRaporKabulNo({barkod},{tc},1,'tr')")
                       time.sleep(14)
                       old_file = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", "Enabiz-PCRSonuc.pdf")
                       new_file = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", f"{isim}.pdf")
                       os.rename(old_file, new_file)
                       time.sleep(8)
                       driver.get('https://enabiz.gov.tr/PcrTestSonuc/Index')
                       time.sleep(3)
                       driver.execute_script(f"GetPcrRaporKabulNo({barkod},{tc},1,'en')")
                       time.sleep(14)
                       old_file1 = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", "Enabiz-PCRSonuc.pdf")
                       new_file1 = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", f"{isim} ENG.pdf")
                       os.rename(old_file1, new_file1)
                       time.sleep(8)
                       deneme += 1
                    except:
                        print(f"{isim} Hatalı")
                        deneme += 1