from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from openpyxl import Workbook
import csv

browser = webdriver.Firefox()
browser.get('https://kontragent.pro/')

inns = []

with open('Список_участников_оборота_d30de03765a1de19d7c6b4b0392866a0.csv',encoding="utf-8") as file:
    read_file = csv.reader(file)
    for row in read_file:
        string = row[0].split(';')[0].split('Свердловская')[0].strip().strip("'")
        inns.append(string)

inns_org = []
names_org = []
phones_org = []
emails_org= []

for inn in inns[:374]:
    if len(inn) < 11:
        sleep(3)
        browser.find_element(By.XPATH, '/html/body/header/div[2]/div/div/form/input').send_keys(inn)
        sleep(3)
        button_search = browser.find_element(By.XPATH, '/html/body/header/div[2]/div/div/form/a').click()
       
        try:
            inns_org.append(browser.find_element(By.XPATH, '/html/body/main/div[4]/section[2]/table/tbody/tr[2]/td').text)
        except:
            inns_org.append('-')
        
        try:
            names_org.append(browser.find_element(By.XPATH, '/html/body/main/div[4]/section[1]/table[1]/tbody/tr[1]/td').text)
        except:
            names_org.append('-')
        
        try:
            phone = browser.find_element(By.XPATH,'/html/body/main/div[4]/section[3]/table/tbody/tr[2]').text
            phones_org.append(phone)         
        except:
            phones_org.append('-')
        
        try:
            emails_org.append(browser.find_element(By.XPATH, '/html/body/main/div[4]/section[3]/table/tbody/tr[3]/td/a').text)
        except:
            try:
                emails_org.append(browser.find_element(By.XPATH, '/html/body/main/div[4]/section[3]/table/tbody/tr[4]/td/a').text)
            except:
                emails_org.append('-')      

work_book = Workbook()
work_book.active
work_sheet_2 = work_book.create_sheet('шины',0)
work_sheet_2.append(['ИНН', 'Название','Телефон', 'Почта'])

row1 = 1
row2 = 1
row3 = 1
row4 = 1

for inn in inns_org:
    row1 += 1
    work_sheet_2['A' + str(row1)] = inn
    
for name_organization in names_org:
    row2 += 1
    work_sheet_2['B' + str(row2)] = name_organization
    
for telephone in phones_org:
    row3 += 1
    work_sheet_2['C' + str(row3)] = telephone
    
for email in emails_org:
    row4 += 1
    work_sheet_2['D' + str(row4)] = email

work_book.save('маркировка шин (ЮЛ из ЧЗ).xlsx')