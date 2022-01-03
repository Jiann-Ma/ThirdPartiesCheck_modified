#!/usr/bin/python
# coding:utf-8

import os
import logging
import datetime
import warnings
warnings.filterwarnings('ignore')
import tkinter as tk
from tkinter import messagebox
from tkinter import Checkbutton, IntVar

class MyApp(object):
    """
    define the GUI interface
    """
    def __init__(self):
        '''
        set initial UI 
        '''
        self.set_log()
        self.root = tk.Tk()
        self.root.title("RPA 司法院查詢")
        self.root.geometry('1000x400')
        self.canvas = tk.Canvas(self.root, height=280, width=500)
        self.canvas.pack(side='top')
        self.setup_ui()
        self.create_icon()
        self.step1_status = False 


    def set_log(self):
        '''
        set log
        '''
        if not os.path.exists('./log'):
            os.mkdir('./log')
        log_time = datetime.datetime.now().strftime('%Y%m%d')
        log_name = f'log/RPA_{log_time}.log'
        self.logger = logging.getLogger(log_name)
        self.logger.setLevel('INFO')
        BASIC_FORMAT = "%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s"
        DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
        formatter = logging.Formatter(BASIC_FORMAT, DATE_FORMAT)
        chlr = logging.StreamHandler() # 輸出到控制台的handler
        chlr.setLevel('INFO')
        fhlr = logging.FileHandler(log_name) # 輸出到文件的handler
        fhlr.setFormatter(formatter)
        self.logger.addHandler(chlr)
        self.logger.addHandler(fhlr)

    def setup_ui(self):
        """
        setup UI interface
        """
        #說明: 指定執行該程式需要具備的web driver
        text1 = '''請確定指定資料夾有檔案：chromedriver.exe、companies.csv'''
        self.label_text1 = tk.Label(self.root, text=text1, bg='white', font=('微軟正黑體', 12), width=90, height=2)
        entry_var = tk.StringVar()
        entry_var.set(r'我是範例 => C:\Users\ESB18507\Github\RPA_CJ_00_001')
        
        #TODO 欄位定義
        #說明: 指定執行該程式需要具備的檔案       
        self.label_file_path = tk.Label(self.root, text='公司列表：')
        self.input_file_path = tk.Entry(self.root, width=100, textvariable=entry_var)    
        self.check_button_value = IntVar()
        self.check_button = Checkbutton(text = "測試模式", variable = self.check_button_value, \
                                        onvalue = 1, offvalue = 0, height=5, \
                                        width = 20)
        #定義按鈕
        self.run_button = tk.Button(self.root, command=self.run, text="執行", width=20)
        self.quit_button = tk.Button(self.root, command=self.quit, text="退出程式", width=20)

    def create_icon(self):
        '''
        create icon
        '''
        import base64
        from icon import iconImg
        with open('tmp.ico', 'wb+') as tmpIcon:
            tmpIcon.write(base64.b64decode(iconImg))
        self.root.iconbitmap('tmp.ico')
        os.remove('tmp.ico')

    def gui_arrang(self):
        """
        setup position of UI
        """
        self.label_text1.place(x=60, y=20)

        #TODO 定義欄位位置
        self.label_file_path.place(x=60, y=80)
        self.input_file_path.place(x=170, y=80)

        self.run_button.place(x=290, y=330)
        self.quit_button.place(x=590, y=330)
        self.check_button.place(x=250, y=220)

    def check(self):
        """
        check the input of gui interface
        return:
            True
            False
        """

        #TODO 提取參數以及進行驗證
        self.path = self.file_path.get().strip()
  
        try:
            #TODO 資料檢查
            if len(self.file_path) == 0:
                messagebox.showinfo(title='System Alert', message='公司列表之路徑不得為空!')
                self.logger.warning('公司列表之路徑為空值!')
                return False
            
        except FileNotFoundError:
            messagebox.showinfo(title='System Alert', message='檔案路徑有問題！')
            self.logger.warning(FileNotFoundError, exc_info=True)
            return False

        except ImportError:
            messagebox.showinfo(title='System Alert', message='無法讀取檔案(ImportError)！')
            self.logger.error(ImportError, exc_info=True)
            return False

        except OSError:
            messagebox.showinfo(title='System Alert', message=f'無法讀取檔案！檔案路徑異常或不存在:{self.file_path}')
            self.logger.error(OSError, exc_info=True)
            return False

        except Exception as error:
            messagebox.showinfo(title='System Alert', message=error)
            self.logger.error(error, exc_info=True)
            return False

        return True

    def run(self):
        """
        when you click the button of step1, it'll execute Run Method
        """
        
        start_time = datetime.datetime.now()
        try:
            state = False
            self.logger.info("測試模式:{}".format(state))
            if self.check() == True:
                return 
                self.main()
            else:
                self.logger.warning('檢查不通過！')

        except Exception as error:
            self.logger.error(error, exc_info=True)
            messagebox.showinfo(title='System Alert', message='執行異常！請洽資安管理處')
            raise error

        finally:
            end_time = datetime.datetime.now()
            execution_time = (end_time-start_time).seconds
            execution_time_format = str(datetime.timedelta(seconds=execution_time))
            self.logger.info('Total Execution time:{}'.format(execution_time_format))
            messagebox.showinfo(title='System Alert', message='執行結束！')

    def quit(self):
        """
        when you click the button of quit, it'll execute
        """
        self.root.destroy()

def main():
    """
    main function for MyApp
    """
    # initial
    app = MyApp()
    # arrage gui
    app.gui_arrang()
    # run tkinter
    tk.mainloop()

if __name__ == '__main__':
    main()
