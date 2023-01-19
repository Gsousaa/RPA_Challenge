from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from pathlib import Path
import os 
import pandas as pd
import time
import datetime

options = webdriver.ChromeOptions()
prefs = {"download.default_directory" : str(Path().resolve())+"\\rpa_challenge"}
options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)

driver.get("https://www.rpachallenge.com/")
download = driver.find_element(By.XPATH,'//*[@class=" col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center"]')
download.click()
time.sleep(1)

input = pd.read_excel(Path("rpa_challenge/challenge.xlsx"))
start = driver.find_element(By.XPATH,'//*[@class="waves-effect col s12 m12 l12 btn-large uiColorButton"]')
start.click()
for index, row in input.iterrows():
    
    first_name = driver.find_element(By.XPATH,'//*[@ng-reflect-name="labelFirstName"]')
    first_name.send_keys(row[0])
    company_name = driver.find_element(By.XPATH,'//*[@ng-reflect-name="labelCompanyName"]')
    company_name.send_keys(row[2])
    phone_number = driver.find_element(By.XPATH,'//*[@ng-reflect-name="labelPhone"]')
    phone_number.send_keys(row[6])
    last_name = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]')
    last_name.send_keys(row[1])
    email = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]')
    email.send_keys(row[5])
    role_company = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]')
    role_company.send_keys(row[3])
    address = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]')
    address.send_keys(row[4])
    submit = driver.find_element(By.XPATH,'//*[@class="btn uiColorButton"]')
    submit.click()
    
time.sleep(5)  
os.remove(str(Path().resolve())+"\\rpa_challenge\\challenge.xlsx")