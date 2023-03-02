#Truesigt_export 
from selenium  import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from time import sleep
from selenium.webdriver import ActionChains







class InstaBot:
    def __init__(self,username,password):  
        #webdriver
        serv_obj = service= Service('geckodriver.exe')
        ops=webdriver.FirefoxOptions()
        ops.headless=True        
        ops.add_argument("--disable-notifications")
        self.driver = webdriver.Firefox(service=serv_obj,options=ops)
        #launch
        self.driver.get("https://presentation.egyptpost.local")
        sleep(2)
        #accept the alert
        self.driver.implicitly_wait(60)
        alert_window= self.driver.switch_to.alert                    #to get the alert message to var
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
        self.driver.find_element(By.XPATH, "//a[normalize-space()='ITSM not assgined']").click()
        sleep(1)
        #time
        self.driver.find_element(By.XPATH, "//a[@class='fi menuIcon fi-action-menu dropdown-toggle']").click()
        sleep(1)
        #time 1hour
        self.driver.find_element(By.XPATH, "//a[@id='1']").click()
        sleep(1)


    def check_Incident_ID(self):
        try:
            #refersh
            self.driver.find_element(By.XPATH, "//a[@id='eventListRefreshBtn']").click()
            sleep(3)
            #total Events
            Total_Events = self.driver.find_element(By.XPATH, "//span[@class='title-display ng-binding']")
        except:
            pass


        if Total_Events.text !=  "Total Events: 0":
            print(Total_Events.text)
             #select all 
            try:
                select_all = self.driver.find_element(By.XPATH, "//span[@class='jqx-checkbox-check-indeterminate']")
                select_all.click()
                #print("select all ",select_all)
                sleep(1)
            except:
                #print("into except")
                select_all = self.driver.find_element(By.XPATH, "//div[@id='row0eventGrid']/div[1]/div[1]")
                select_all.click()   
                pass



            check_Box = self.driver.find_element(By.XPATH, "/html/body/article/div/ng-include[2]/div/div[3]/div[1]/section[1]/div[1]/div/div/div/div[4]/div[1]/div/div[1]/div/div/div")
            if check_Box.is_selected():
                #print("check box passed")
                pass
            else:
                check_Box.click()
                #print("from condition")
            #launch_remote_Action_id
            sleep(1)
            try:
                lanuch = self.driver.find_element(By.XPATH, "//a[@id='launch_remote_Action_id']")
                lanuch.click()
                sleep(1)
                #Trigger Remedy
                self.driver.find_element(By.XPATH, "//li[@id='6']//div[@class='draggable jqx-rc-all jqx-tree-item jqx-item']").click()
                sleep(1)
                #Launch
                self.driver.find_element(By.XPATH, "//button[@id='lunchActionButton']").click()
                sleep(1)
                #write rbsd
                self.driver.find_element(By.XPATH, "//input[@id='DestinationInput']").send_keys("ibrsd")
                sleep(4)
                #excute
                excute = self.driver.find_element(By.XPATH, "//button[@id='executeActionID']")
                excute.click()
                sleep(2)
                if excute.is_enabled:
                    excute.click()

                sleep(15)
                ok_button = self.driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[3]/button")
                if excute.is_enabled:
                    ok_button.click()
                else:
                    pass

            except:
                act = ActionChains(self.driver)
                act.key_down(Keys.ESCAPE,Keys.ESCAPE).perform()
                pass
#pyinstaller --onefile --add-data "wavz.png;." -w -i noc.ico filter.py




        

          

# user_id=input('Enter Your Username:')
# password=getpass('Enter Password:')
# #user_id=""
#password=""
mybot=InstaBot("w_abdelrahman.ataa", "a1591997A!")
#mybot=InstaBot("w_a.aly", "asd@123#456")
#ITSM not assgined
mybot.choose_filter()

while True:
    print("waiting 30sc")
    sleep(30)
    try:
        mybot.check_Incident_ID()
    except:
        pass

