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
first = 0
#read google-sheets
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSINLlSv4NcCszvA5XOPsuYCxZEk9_tBnhgLvyDkcG73QgFObITFtaZRQ492wlS53NPBlQi0AfPHMVh/pub?gid=1326092367&single=true&output=csv"
code = pd.read_csv(url)
rename_data = []
for i in range(0,code.shape[0]):
    if code["證券代號"][i] == "#REF!":
        count+=1
    else:
        if first == 0:
            first+=1
            rename_data.append({"original_filename": "K_Chart.csv", "new_filename": str(code["證券代號"][i])+".csv"})
        else:
            rename_data.append({"original_filename": "K_Chart ("+str(i-count)+").csv", "new_filename": str(code["證券代號"][i])+".csv"})



df = pd.DataFrame(rename_data)

# 指定輸出的檔案路徑
codes_dir = "codes"
os.makedirs(codes_dir, exist_ok=True)  # 如果資料夾不存在，則建立
output_file = os.path.join(codes_dir, "rename_mapping.csv")

# 儲存為 CSV 檔案
try:
    df.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"檔案已成功儲存至：{output_file}")
except Exception as e:
    print(f"儲存檔案失敗：{e}")














