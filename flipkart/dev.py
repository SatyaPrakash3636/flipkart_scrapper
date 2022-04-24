from bs4 import BeautifulSoup
from selenium import webdriver
# from selenium.webdriver.support.ui import WebDriverWait
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.support import expected_conditions as EC
from flipkart import constant as const
import datetime
import time
import os


class Flipkart(webdriver.Chrome):

    def __init__(self, target_dir=None):
        self.target_dir = target_dir
        self.service = Service(executable_path= "F:\Python_learning\webscraping\chromedriver_win32\chromedriver.exe")
        chrome_options = Options()
        # chrome_options.headless = True
        if self.target_dir == None:
            self.target_dir = os.getcwd() + "\defalut_dir"
            if not os.path.exists(self.target_dir):
                os.makedirs(self.target_dir)

        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": f"{self.target_dir}",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False
            })
        super().__init__(service=self.service, options=chrome_options)

        self.implicitly_wait(8)
        self.maximize_window()
        

    def landing_page(self):
        self.get(const.BASE_URL)

    def login_page(self):
        login_button = self.find_element(By.CLASS_NAME, 'fGVtXS')
        login_button.click()
        user_name = self.find_element(
            By.CSS_SELECTOR, 'input[placeholder="Enter your username OR Phone Number OR Email"]'
        )
        user_name.send_keys(const.LOGIN_USER)
        time.sleep(1)
        next_button = self.find_element(
            By.XPATH, '//*[@id="app"]/div[3]/section/section/div/div[2]/button/span'
        )
        next_button.click()
        enter_pass = self.find_element(
            By.XPATH, '//*[@id="app"]/div[3]/section/section/div/div[1]/form/div[2]/div/input'
        )
        enter_pass.send_keys(const.LOGIN_PASS)
        login_button = self.find_element(
            By.XPATH, '//*[@id="app"]/div[3]/section/section/div/div[2]/button/span'
        )
        login_button.click()
        time.sleep(5)

    def get_all_seller(self):
        self.landing_page()
        self.login_page()
        html_data = self.execute_script("return document.documentElement.outerHTML")
        logout_element = self.find_element(By.XPATH, '//*[@id="blinx-wrapper-0"]/div[1]/div[1]/div/div[2]/div[1]')
        logout_element.click()
        sel_soup = BeautifulSoup(html_data, 'lxml')
        sellers_html = sel_soup.find('ul', class_="fk-list")
        sellers_list = sellers_html.find_all('li', class_="cf-list")
        sellers_details = []
        comp_id = 0
        for seller in sellers_list:
            comp_id += 1
            seller_name = seller.find('div', class_="col-md-8").text.strip()
            seller_comp = seller.find('div', class_="col-xs-12").text.strip()
            seller_info = [comp_id, seller_name, seller_comp]
            sellers_details.append(seller_info)
        return sellers_details

    def select_seller(self, i):
        global seller_id
        seller_id = i
        seller = self.find_element(
            By.XPATH, f'//*[@id="blinx-wrapper-2"]/div/ul/li[{seller_id}]/label/div/div[2]/div[1]/div[1]'
        )
        seller.click()
        access = self.find_element(
            By.XPATH, '//*[@id="blinx-wrapper-0"]/div[2]/div[3]/div/div[3]/button'
        )
        access.click()
        time.sleep(5)

    def earn_more(self, s_name, period="weekly"):
        def close_ad():
            try:
                skip = self.find_element(
                    By.XPATH, '//*[@id="react-joyride:0"]/div/div/div[1]/div[2]/button[1]'
                )
                skip.click()
            except:
                pass
            try:
                self.find_element(By.XPATH, '//*[@id="desktopBannerWrapped"]/div/div[3]/div[1]/button[1]').click() # Dont Allow button
                #self.find_element(By.XPATH, '//*[@id="optInText"]') # Allow Button
            except:
                pass
            return

        def click_download():
            try:
                download_element = self.find_element(By.XPATH, '//*[@id="sub-app-container"]/section/section[1]/section/div[2]/div[2]/div/button')
                download_element.click()
                print("Clicked Download")
                time.sleep(10)
            except:
                pass

        close_ad()
        self.get('https://seller.flipkart.com/index.html#dashboard/growth/earn-more')
        time.sleep(2)
        close_ad()
        if period == "weekly":
            pass
        elif period == "monthly":
            monthly_element = self.find_element(By.XPATH, '//*[@id="sub-app-container"]/section/section[1]/section/div[1]/div/div/div[3]/div')
            monthly_element.click()
            close_ad()
            time.sleep(1)
        elif period == 'latest':
            # todays_date = date.today()
            # startDate = enddate = todays_date - timedelta(days=2)
            # time.sleep(4)
            latest_element = self.find_element(By.XPATH, '//*[@id="sub-app-container"]/section/section[1]/section/div[1]/div/div/div[1]/div')
            latest_element.click()
            time.sleep(1)
            # self.get(f'https://seller.flipkart.com/index.html#dashboard/growth/earn-more?selected=latest&startDate={startDate}&endDate={enddate}')
            close_ad()
        click_download()
        while True:
            target_file = self.target_dir + "\\" + "earn_more_report.xlsx"
            if not os.path.exists(target_file):
                print("In while loop")
                click_download()
                continue
            else:
                break
        time_now = datetime.datetime.now()
        time_formatted = time_now.strftime("%Y-%m-%d_%H-%M")
        rename_to = self.target_dir + "\\" + s_name + "-" + period + "-" + time_formatted + ".xlsx"
        os.rename(target_file, rename_to)
        return rename_to
