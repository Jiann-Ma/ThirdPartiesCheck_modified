# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 09:09:08 2021

@author: ESB20914
"""

# selenium libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException

from tkinter import messagebox
import pandas as pd
import time, os


class judicialRecordsCheck:
    
    def __init__(self, file_path, state, background=False):
        self.file_path = file_path
        self.debug = state
        self.background = background
        self.set_driver()
        
    def set_driver(self):
        # set chrome driver path
        executable_path ='../chromedriver.exe'
        # chrome_options
        chrome_options = Options()
        if self.background:
            chrome_options.add_argument('--headless')
        # avoid google bug
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-default-apps")
        chrome_options.add_argument("--disable-dev-shm-usage")

        # get the driver
        try:
            driver = webdriver.Chrome(executable_path=executable_path, 
                                    chrome_options=chrome_options)
        except WebDriverException as error:
            self.logger.error(error, exc_info=True)
            infomation = """- 此錯誤為 chromedrive.exe 版本與 chrome 版本不合導致，chrome 會自動更新
                            - 請先查看 chrome 的版本之後
                            - 再到 https://chromedriver.chromium.org/downloads 進行下載對應版本
                            - 下載 win32 版本即可
                            - 將檔案解壓縮，把 chromedriver.exe 覆蓋指定的資料夾路徑即可
                         """
            messagebox.showinfo(title='System Alert', message=infomation)
            raise error
        # get url
        driver.get("https://law.judicial.gov.tw/FJUD/default.aspx")
        driver.maximize_window()
        self.driver = driver

        
    def main(self):
        # Reference: https://blog.impochun.com/excel-big5-utf8-issue/
        df_input = pd.read_csv('..//companies.csv', encoding= 'utf-8-sig', converters={'companies':str})
        print(f'[INFO] 已成功讀取資料，筆數: {len(df_input)}，內容如以下：')
        print(df_input['companies']) 
    

        for row in df_input.index:
        
            time.sleep(2)
            forVerify = df_input.loc[row, 'companies']
        
            # 加上編號
            forVerifyNo = f'{row}' + ". " + forVerify
        
            # 輸入要檢查的項目
            needsCheck = WebDriverWait(self.driver, 2).until(EC.visibility_of_element_located((By.ID, 'txtKW')))
            needsCheck.send_keys(forVerify)
    
            time.sleep(2)
        
            #按下送出查詢
            self.driver.find_element(By.XPATH, '//*[@id="btnSimpleQry"]').click()
            time.sleep(2)
        
            #截圖
            directory_time = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    
            # 印出當前目錄以確認
            print(os.getcwd())
            try:
                file_path = os.getcwd() + '\\' + directory_time + '\\'
                if not os.path.exists(file_path):
                    os.makedirs(file_path)
                    print('[INFO] 目錄新建成功：%s' % file_path)
                else:
                    print('[INFO] 目錄已存在！')
            except BaseException as msg:
                print('[INFO] 目錄新建失敗：%s' % msg)
    
            self.driver.save_screenshot(file_path + '\\' + f'{forVerifyNo}.png')
            print(f'[INFO] 拍完第{row}筆的照片了！')
            time.sleep(2)
        
            #點下「判決書查詢」，回到查詢頁面
            self.driver.find_element(By.XPATH, '//a[@href="/FJUD/default.aspx"]').click()
        
            #清除輸入的字(cookies)
            self.driver.delete_all_cookies()

            print(f'[INFO] 現在做完第{row}筆')

        print(f"[INFO] 總共{len(df_input)}筆，全部完成了！")
        self.driver.close()
