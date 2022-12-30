#GLOBALS
USERLOGIN = ""
USERPASSWORD = ""


import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
import datetime

start_time = str(datetime.datetime.now())

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

driver.get("https://na2.exo5.com/assets/")

username = driver.find_element(By.ID, "id_username")
username.send_keys(USERLOGIN)

password = driver.find_element(By.ID, "id_password")
password.send_keys(USERPASSWORD)

login = driver.find_element(By.ID, "submit_login")
login.click()


# Set second element in dropdown menu to 1000, select that option
size = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div/div[1]/div/div[1]/div[2]/div/div[1]/div/div[3]/select/option[2]")
driver.execute_script("arguments[0].setAttribute('value', '1000')", size)
size_list = Select(driver.find_element(By.ID, "page_size_select"))
size_list.select_by_value("1000")

# wait 20 second while page loads all elements
for i in range(20):
    print(20-i)
    time.sleep(1)

device_list = driver.find_element(By.CSS_SELECTOR, "#assets_paginator tbody")

devices = device_list.find_elements(By.CSS_SELECTOR, ".system a")

print(devices)

original_window = driver.current_window_handle

machines = []

for item in devices:
    data = {"user":"", "password":"", "encryption": "", "bitlocker":"", "identifier":"", "model":"", "serial":""}
    href = item.get_attribute("href")
    print("Scraping data from: " + href)
    driver.switch_to.new_window('window')
    time.sleep(1)
    driver.get(href)
    time.sleep(2)
    try:
        username = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div/div[1]/div/table/tbody/tr[2]/td[2]").text
        data["user"] = username
        print(username)
    except:
        print("Error getting username")
    driver.get(href + "remotekill/2")
    time.sleep(2)

    try:
        password = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div[2]/div/div[2]/div/div[2]/div[3]/div/div[2]/div/div[1]/span").text
        data["password"] = password
        print(password)
    except:
        print("Error getting password")
    driver.get(href + "diskencryption")
    time.sleep(2)

    try:
        encryption = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div[2]/div/div[1]/div/div[1]/div/div/div[1]/div[1]/div/div/div[3]/div/div/div/div[3]/div[1]/div/div/span").text
        data["encryption"] = encryption
        print(encryption)
    except:
        print("Error getting encryption")
    try:
        bitlocker = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div[2]/div/div[1]/div/div[1]/div/div/div[5]/div/div[1]/span[2]").text
        data["bitlocker"] = bitlocker
        print(bitlocker)
        identifier = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div/div/div[3]/div/div/div[2]/div/div[1]/div/div[1]/div/div/div[5]/div/div[1]/span[4]").text
        data["identifier"] = identifier
        print(identifier)
    except:
        print("Error getting bitlocker data (probably unencrypted)")
    driver.get(href + "hardware")
    time.sleep(2)
    try:
        model = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr[2]/td[2]").text
        data["model"] = model
        print(model)
        serial = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr[3]/td[2]").text
        data["serial"] = serial
        print(serial)
        print(data)
    except:
        print("Error getting hardware data")
    machines.append(data)
    driver.close()
    driver.switch_to.window(original_window)

driver.quit()

print('''
///////////////////////////////
/// DATA SCRAPING FINISHED! ///
///////////////////////////////                                                                                                
''')

end_time = str(datetime.datetime.now())

print("Creating xlsx file")
workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "USERNAME"
sheet["B1"] = "BOOT SECTOR PASSWORD"
sheet["C1"] = "ENCRYPTION"
sheet["D1"] = "BITLOCKER KEY"
sheet["E1"] = "BITLOCKER IDENTIFIER"
sheet["F1"] = "MODEL NAME"
sheet["G1"] = "SERIAL NUMBER"
sheet["H1"] = "SCRAPING START"
sheet["I1"] = "SCRAPING END"
sheet["H2"] = start_time
sheet["I2"] = end_time


for i in range(2, len(machines)+2):
    sheet[("A" + str(i))] = machines[i-2]["user"]
    sheet[("B" + str(i))] = machines[i-2]["password"]
    sheet[("C" + str(i))] = machines[i-2]["encryption"]
    sheet[("D" + str(i))] = machines[i-2]["bitlocker"]
    sheet[("E" + str(i))] = machines[i-2]["identifier"]
    sheet[("F" + str(i))] = machines[i-2]["model"]
    sheet[("G" + str(i))] = machines[i-2]["serial"]

workbook.save(filename = f'exodata{end_time[:-16]}.xlsx')

print("Excel file created")
print("Exo scraper finished!")