from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import chromedriver_autoinstaller
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta


from pathlib import Path
import pandas as pd
import time
import os
from os.path import exists
import shutil
import csv

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
preferences = {
    "download.default_directory": downloadDir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0,
    "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    "profile.default_content_setting_values.automatic_downloads": 1,}
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

count = 0
#read google-sheets
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSINLlSv4NcCszvA5XOPsuYCxZEk9_tBnhgLvyDkcG73QgFObITFtaZRQ492wlS53NPBlQi0AfPHMVh/pub?gid=1326092367&single=true&output=csv"
code = pd.read_csv(url)

for i in range(0,code.shape[0]):
    try:
        driver.get('https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID='+str(int(code["證券代號"][i]))+'&CHT_CAT=WEEK')
        if count == 0:
            print("a")
            time.sleep(1)
            count+=1
            try:
                # 等待遮擋按鈕出現，最多等 10 秒
                interstitial_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ats-interstitial-button"))
                )
                time.sleep(1)  # 稍作停留確保穩定
                interstitial_button.click()
                print("pass window")
            except Exception as e:
                print("undetect window...keep going", e)

        month = list(code["撥券日期(上市、上櫃日期)"])

        #tiding
        month_pass = ""
        for k in range(0,code.shape[0]):
            month_pass = month_pass+str(month[k])+"/"
        month = month_pass.split("/")
        #month = [2024,12,4,2024,11,3.....]   
        print(month)

        s = driver.find_element(By.ID, "edtSTART_TIME")
        for z in range(len(month)):
            try:
                int(month[z])
            except:
                month.pop(z)


        while True:
            try:
                if int(month[1+i*3]) == 1 or int(month[1+i*3]) == 2:
                    month[1+i*3] = int(month[1+i*3])+10
                    month[0+i*3] = int(month[0+i*3])-1
                    break
                else:
                    month[1+i*3] = int(month[1+i*3])-2
                    break
            except:
                print(month[1+i*3])
                print("AA?")
                if   > code.shape[0]:
                    break

      
        s.click()
        time.sleep(5)
        s.send_keys(Keys.LEFT)
        s.send_keys(Keys.LEFT)

        #輸入日期
        for g in range(0,2):
            if g == 1:
                s.send_keys(Keys.TAB)
            code_send = month[(1-g)+i*3]
            print(code_send)
            s.send_keys(code_send)
            time.sleep(1)





      
        element = driver.find_element("xpath", "//input[@value='查詢']")
        element.click()
        time.sleep(1)
        
        element = driver.find_element("xpath", "//input[@value='XLS']")        
        element.click()
        time.sleep(1)






        



        #driver.get_screenshot_as_file("page.png")
        latestDownloadedFileName = getDownLoadedFileName() 
        time.sleep(1)
        #driver.get_screenshot_as_file("page1.png")
        getDownLoadedFileNameClose()
        DownloadedFilename=''.join(latestDownloadedFileName).encode().decode("utf-8")
          
        if DownloadedFilename != "OTC.csv":
            # Copy the file to "OTC.csv"
            shutil.copy(DownloadedFilename, "OTC.csv")
            print(f"File '{DownloadedFilename}' copied to 'OTC.csv'.")
            print("Download completed...",downloadDir+'OTC.csv')
            downloaded_files = os.listdir(downloadDir)
            print(downloaded_files) 
    except:                   
        continue



