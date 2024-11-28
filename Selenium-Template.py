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
import requests
from bs4 import BeautifulSoup


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

count = 0
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
        if count == 0:
            time.sleep(5)
            count+=1
            try:
                # 等待遮擋按鈕出現，最多等 10 秒
                interstitial_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ats-interstitial-button"))
                )
                time.sleep(2)  # 稍作停留確保穩定
                interstitial_button.click()
                print("成功跳過視窗")
            except Exception as e:
                print("未偵測到跳視窗或按鈕，繼續執行：", e)

        month = list(code["撥券日期(上市、上櫃日期)"])

        #tiding
        month_pass = ""
        for k in range(0,code.shape[0]):
            month_pass = month_pass+str(month[k])+"/"
        month = month_pass.split("/")
        #month = [2024,12,4,2024,11,3.....]   
        print(month)

        s = driver.find_element(By.ID, "edtSTART_TIME")



      
      
        if int(month[1+i*3]) == 1 or int(month[1+i*3]) == 2:
            month[1+i*3] = int(month[1+i*3])+10
            month[0+i*3] = int(month[0+i*3])-1
        else:
            month[1+i*3] = int(month[1+i*3])-2

      
        s.click()
        time.sleep(5)
        s.send_keys(Keys.LEFT)
        s.send_keys(Keys.LEFT)

        #輸入日期
        for g in range(0,3):
            s.send_keys(Keys.LEFT)
            s.send_keys(Keys.LEFT)
            for i in range(0,g):
                s.send_keys(Keys.TAB)
            code_send = month[(1-g)+i*3]
            print(code_send)
            s.send_keys(2022)
            time.sleep(5)





      
        element = driver.find_element("xpath", "//input[@value='查詢']")
        element.click()
        time.sleep(5)
        
        element = driver.find_element("xpath", "//input[@value='XLS']")        
        element.click()
        time.sleep(5)






        """stock_id = str(code["證券代號"][i])
        end_date = datetime.now()  # 今天
        start_date = end_date - timedelta(days=60)
        # 模擬瀏覽器的 HTTP 標頭
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
            "Accept-Language": "zh-TW",
        }
        
        # 目標 URL
        base_url = "https://goodinfo.tw/tw/ShowK_Chart.asp"
        
        # 發送 GET 請求並獲取網頁
        params = {
            "STOCK_ID": stock_id,
            "CHT_CAT": "WEEK",
        }
        response = requests.get(base_url, headers=headers, params=params)
        
        if response.status_code == 200:
            # 使用 BeautifulSoup 解析網頁內容
            soup = BeautifulSoup(response.text, "html.parser")
        
            # 嘗試找到下載 XLS 的連結
            download_link = None
            for a_tag in soup.find_all("a"):
                if "xls" in a_tag.get("href", ""):
                    download_link = a_tag["href"]
                    break
        
            if download_link:
                # 完整的 XLS 下載連結
                xls_url = f"https://goodinfo.tw{download_link}"
        
                # 發送請求下載 XLS 檔案
                xls_response = requests.get(xls_url, headers=headers)
        
                if xls_response.status_code == 200:
                    # 儲存檔案到本地
                    file_name = f"{stock_id}_data.xls"
                    with open(file_name, "wb") as f:
                        f.write(xls_response.content)
                    print(f"成功下載 Excel 檔案，儲存為：{file_name}")"""



        #driver.get_screenshot_as_file("page.png")
        latestDownloadedFileName = getDownLoadedFileName() 
        time.sleep(5)
        #driver.get_screenshot_as_file("page1.png")
        getDownLoadedFileNameClose()
        DownloadedFilename=''.join(latestDownloadedFileName).encode().decode("utf-8")
          
        if DownloadedFilename != "OTC.csv":
            # Copy the file to "OTC.csv"
            shutil.copy(DownloadedFilename, "OTC.csv")
            print(f"File '{DownloadedFilename}' copied to 'OTC.csv'.")
            print("Download completed...",downloadDir+'OTC.csv')

"""b = []
end = []
count = 0
for i in range(len(month)):
          b.append(month[i])
          count+=1
          if count == 3:
                    b.append(code["證券代號"][i])
                    end.append(b)
                    b = []
                    count = 0"""



"""# 設定檔案資料夾的路徑
folder_path = 'file'

# 遍歷資料夾中的檔案並更名
for idx, file_name in enumerate(os.listdir(folder_path)):
    if file_name.startswith("K_Chart") and file_name.endswith(".xls"):
        # 確保 list 中有足夠的資料來進行更名
        if idx < len(end):
            # 提取資料列表中的資訊
            year, month, day, code = end[idx]
            # 設定新檔案名稱
            new_name = f"{code}-{month}-{day}.xls"
            old_path = os.path.join(folder_path, file_name)
            new_path = os.path.join(folder_path, new_name)
            
            # 更改檔案名稱
            os.rename(old_path, new_path)
            print(f"Renamed '{file_name}' to '{new_name}'")"""

