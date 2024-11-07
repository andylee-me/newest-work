from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import chromedriver_autoinstaller

from pathlib import Path
import pandas as pd
import time
import os
from os.path import exists
import shutil
import csv
import keys

# The following 3 lines are for ubuntu only. If windows, please comments then to work well..
from pyvirtualdisplay import Display
display = Display(visible=0, size=(800, 800))  
display.start()

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path


def getDownLoadedFileNameClose():
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
def getDownLoadedFileName():
    driver.execute_script("window.open()")
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('chrome://downloads')
    #driver.get_screenshot_as_file("page.png")
    return driver.execute_script("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
 
  
downloadDir = f"{os.getcwd()}//"
preferences = {"download.default_directory": downloadDir,
                "download.prompt_for_download": False,
                "directory_upgrade": True,
                "safebrowsing.enabled": True}
chrome_options = webdriver.ChromeOptions()  

chrome_options.add_experimental_option("prefs", preferences)
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument("--window-size=1200,1200")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--no-sandbox")

    
driver = webdriver.Chrome(options = chrome_options)

#driver.get('http://github.com')
#print(driver.title)
#with open('./GitHub_Action_Results.txt', 'w') as f:
#    f.write(f"This was written with a GitHub action {driver.title}")


#read google-sheets
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSINLlSv4NcCszvA5XOPsuYCxZEk9_tBnhgLvyDkcG73QgFObITFtaZRQ492wlS53NPBlQi0AfPHMVh/pub?gid=1326092367&single=true&output=csv"
code = pd.read_csv(url)

for i in range(0,code.shape[0]):
    try:                      
        int(code["證券代號"][i])
    except:                   
        break
    
    if int(code["證券代號"][i]) != "":
        driver.get('https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID='+str(int(code["證券代號"][i]))+'&CHT_CAT=WEEK')
        time.sleep(2)

        month = list(code["撥券日期(上市、上櫃日期)"])

        #tiding
        month_pass = ""
        for k in range(0,code.shape[0]):
            month_pass = month_pass+str(month[k])+"/"
        month = month_pass.split("/")
        month.pop(-1) 
        #month = [2024,12,4,2024,11,3.....]                
                
        s = driver.find_element(by=By.CLASS_NAME, value="edtSTART_TIME")
        if int(month[1+i*3]) == 1 or int(month[1+i*3]) == 2:
            month[1+i*3] = int(month[1+i*3])+10
            month[0+i*3] = int(month[0+i*3])-1
        s.click()

        #輸入日期
        for g in range(0,3):
            if g == 1:
                s.send_keys(Keys.TAB)
            for f in range(0,len(month[g+i*3])):
                code_send = month[g+i*3][f]
                s.send_keys(code_send)
                time.sleep(0.01)
        element = driver.find_element(by=By.CLASS_NAME, value="查詢")
        element.click()
        time.sleep(0.5)
        element = driver.find_element(by=By.CLASS_NAME, value="XLS")
        element.click()
        time.sleep(2)



        #driver.get_screenshot_as_file("page.png")
        latestDownloadedFileName = getDownLoadedFileName() 
        time.sleep(2)
        #driver.get_screenshot_as_file("page1.png")
        getDownLoadedFileNameClose()
        DownloadedFilename=''.join(latestDownloadedFileName).encode().decode("utf-8")
          
        if DownloadedFilename != "OTC.csv":
            # Copy the file to "OTC.csv"
            shutil.copy(DownloadedFilename, "OTC.csv")
            print(f"File '{DownloadedFilename}' copied to 'OTC.csv'.")
            print("Download completed...",downloadDir+'OTC.csv')

  