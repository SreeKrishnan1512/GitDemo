
import pytest
import openpyxl
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service

excelPath="C:/Users/psaik/Downloads/downloadExcel.xlsx"
sheet=openpyxl.load_workbook(excelPath)
ApplePrice=0

active=sheet.active
active.cell(row=3,column=4).value="400"

Dict={}

chromePath="C:/Users/psaik/Desktop/Selenium_Excel_New/chromedriver-win64/chromedriver-win64/chromedriver.exe" 
service_obj= Service(chromePath)
excel_Path="C:/Users/psaik/Downloads/downloadExcel.xlsx"
if not os.path.exists(excel_Path):
    print("File not found!")
else:
    print("File exists!")
driver=webdriver.Chrome(service=service_obj)

driver.get("https://rahulshettyacademy.com/upload-download-test/index.html")

driver.maximize_window()

fruit_name="Apple"

WebDriverWait(driver,10).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR,"#downloadButton"))
    ).click()

#for uploading file we have to check whether type='file' is present in DOM, If its present then it means 
# obviously we need to upload some files

input_file=WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"input[type='file']")))
# Calling .click() on the <input> element opens the file picker dialog, but Selenium cannot interact with native 
# system dialogs (like the file picker). Therefore, you bypass this step by directly sending the file path.
input_file.send_keys(excel_Path)
#In send keys we have to share the file path location

text_capture=WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,"//div[text()='Updated Excel Data Successfully.']")))
txt=text_capture.text
priceColumn= driver.find_element(By.XPATH,"//div[text()='Price']").get_attribute("data-column-id")
actual_price=driver.find_element(By.XPATH,f"//div[text()='{fruit_name}']/parent::div/parent::div/div[@id='cell-{priceColumn}-undefined']").text
print(actual_price)
print(txt)
time.sleep(5)

