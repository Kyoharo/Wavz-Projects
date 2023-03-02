#Truesigt_export 
from selenium  import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from time import sleep
import getpass
import glob
import os
import csv 
import openpyxl






class InstaBot:
    def __init__(self,username,password):  
        #webdriver
        serv_obj = service= Service(r'\\10.199.199.35\soc team\Abdelrahman Ataa\Driver\geckodriver.exe')
        ops = webdriver.FirefoxOptions()
        ops.add_argument("--disable-notifications")
        self.driver = webdriver.Firefox(service=serv_obj,options=ops)
        #launch
        self.driver.get("https://presentation.egyptpost.local")
        sleep(2)
        #accept the alert
        self.driver.implicitly_wait(20)
        alert_window= self.driver.switch_to.alert                    #to get the alert message to var
        print(alert_window.text)                                #print the message 
        #alert_window.send_keys("testtt")                        # send keys to pop message 
        alert_window.accept()                              # accept the alert
              
        #Accept the alert
        self.username= username
        self.password= password
        #username
        self.driver.find_element(By.XPATH, '//input[@id="user_login"]').send_keys(username)
        sleep(1)
        #password
        self.driver.find_element(By.XPATH, '//input[@id="login_user_password"]').send_keys(password)
        sleep(1)
        #login
        self.driver.find_element(By.XPATH, '//button[@id="login-jsp-btn"]').click()
        sleep(1)
        self.driver.maximize_window()


    def choose_filter(self):
        #header-menu-toggl
        sleep(1)
        self.driver.find_element(By.XPATH, '//a[@id="header-menu-toggle"]').click()
        sleep(2)
        #Monitoring
        self.driver.find_element(By.XPATH, "//span[normalize-space()='Monitoring']").click()
        sleep(1)
        #Events
        self.driver.find_element(By.XPATH, "//span[normalize-space()='Events']").click()
        sleep(1)
        #eventsViewAction
        self.driver.find_element(By.XPATH, "//span[@id='eventsViewAction']").click()
        sleep(1)
        #View Saved Filters
        self.driver.find_element(By.XPATH, "//li[@id='eventView.filter.savedFilters']//a[@ng-click='clickHandler();']").click()
        sleep(1)
        #View Saved Filters
        self.driver.find_element(By.XPATH, "//a[normalize-space()='Automtaion filter']").click()
        sleep(1)
        #time
        self.driver.find_element(By.XPATH, "//a[@class='fi menuIcon fi-action-menu dropdown-toggle']").click()
        sleep(1)
        #time 1hour
        self.driver.find_element(By.XPATH, "//a[@id='4']").click()
        sleep(1)


    def export(self):
         #select all 
        self.driver.find_element(By.XPATH, "//span[@class='jqx-checkbox-check-indeterminate']").click()
        sleep(1)
        check_Box = self.driver.find_element(By.XPATH, "/html/body/article/div/ng-include[2]/div/div[3]/div[1]/section[1]/div[1]/div/div/div/div[4]/div[1]/div/div[1]/div/div/div")
        if check_Box.is_selected():
            pass
        else:
            check_Box.click()
        #export
        self.driver.find_element(By.XPATH, "//a[@id='show_export_anchor_id']").click()
        sleep(1)
        #csv
        self.driver.find_element(By.XPATH, "//input[@id='radio_format_csv']").click()
        sleep(1)
        #export button
        self.driver.find_element(By.XPATH, "//button[@ng-disabled='eventExport.data.loading || eventExport.data.loadingFailed']").click()
        

    def lastfile_name(self,folder_path):

        files_path = os.path.join(folder_path, '*')
        files = sorted(glob.iglob(files_path), key=os.path.getctime, reverse=True) 
        mylastfile = files[0] 
        return mylastfile

    def csv_to_excel(self,sheet_path,last_file_var):
        #get the data from csv to list
        csv_data = []
        with open(last_file_var) as file_obj:
            reader = csv.reader(file_obj)
            for row in reader:
                csv_data.append(row)

        # Load the workbook
        wb = openpyxl.load_workbook(sheet_path)
        ws = wb["Sheet1"]
        #write the data to the excel file
        for i in range(1,len(csv_data)):
            ws.append(csv_data[i])
        wb.save(sheet_path)



        

          

# user_id=input('Enter Your Username:')
# password=getpass('Enter Password:')
# #user_id=""
#password=""
Devicename = os.getlogin()
last_file_path=r"C:\Users"+  "\\" + Devicename + "\\" + "Downloads" + "\\"

mybot=InstaBot("W_abdelrahman.ataa", "a1591997A!")
mybot.choose_filter()
mybot.export()
lastt_file = mybot.lastfile_name(last_file_path)
mybot.csv_to_excel(r"\\10.199.199.35\soc team\SOC_Daily Report\2023\January\BMC Automation\BMC_Automation_report.xlsx",lastt_file)


print(last_file_path)