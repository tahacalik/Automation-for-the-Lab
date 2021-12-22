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
    id = (excel_worksheet.cell_value(x,2)) #3. sütun
    barcode = str((excel_worksheet.cell_value(x,3)))[:-2] #4. sütun
    patient = (excel_worksheet.cell_value(x,0)) #1. sütun
    driver.get('https://enabiz.gov.tr/PcrTestSonuc/Index')
    time.sleep(5)
    times = 0
    while times < 1:
                    try:
                       driver.execute_script(f"GetPcrRaporKabulNo({barcode},{id},1,'tr')")
                       time.sleep(14)
                       old_file = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", "Enabiz-PCRSonuc.pdf")
                       new_file = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", f"{patient}.pdf")
                       os.rename(old_file, new_file)
                       time.sleep(8)
                       driver.get('https://enabiz.gov.tr/PcrTestSonuc/Index')
                       time.sleep(3)
                       driver.execute_script(f"GetPcrRaporKabulNo({barcode},{id},1,'en')")
                       time.sleep(14)
                       old_file1 = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", "Enabiz-PCRSonuc.pdf")
                       new_file1 = os.path.join(r"C:\Users\AnatoliaPCR\Downloads", f"{patient} ENG.pdf")
                       os.rename(old_file1, new_file1)
                       time.sleep(8)
                       times += 1
                    except:
                        print(f"{patient} error")
                        times += 1
