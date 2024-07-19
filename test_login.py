# main executable file for PYtest
import datetime
from Locators.login_locators import zenLocators
from Utilities.excel_functions import Excel_function
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pytest


class Test_Zen_portal:

    # booting function used with in every pytest which will run with each pytest
    @pytest.fixture()
    def booting_fun(self):
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)
        yield
        self.driver.close()

    # test will store pass and fail data in excel file after execution
    def test_login(self, booting_fun):
        try:
            # creating excel object of excel_function class
            excel_obj = Excel_function(zenLocators.excel_file, zenLocators.sheet_number)
            row = excel_obj.row_count()
            # explict wait object creation
            wait = WebDriverWait(self.driver, 10)
            actions = ActionChains(self.driver)
            for row in range(2, row + 1):
                # reading data from excel file
                username = excel_obj.read_data(row, 6)
                password = excel_obj.read_data(row, 7)
                self.driver.get(zenLocators.url)
                Username = wait.until(EC.presence_of_element_located((By.XPATH, zenLocators().username)))
                Password = wait.until(EC.presence_of_element_located((By.XPATH, zenLocators().password)))
                Login = wait.until(EC.presence_of_element_located((By.XPATH, zenLocators().submit_button)))
                Username.send_keys(username)
                Password.send_keys(password)
                Login.click()
                # validation for checking and writing status of testcase in excel file
                if zenLocators().dashboard_url in self.driver.current_url:
                    print("SUCCESS : Login successful")
                    excel_obj.write_data(row, 8, zenLocators().pass_data)
                    excel_obj.write_data(row, 9, datetime.datetime.now())
                    human_image = self.driver.find_element(By.XPATH, value=zenLocators().drop_down)
                    human_image.click()
                    log_out = self.driver.find_element(By.XPATH, value=zenLocators().log_out)
                    log_out.click()

                elif zenLocators.url in self.driver.current_url:
                    print("ERROR : Login unsuccessful")
                    excel_obj.write_data(row, 8, zenLocators().fail_data)
                    excel_obj.write_data(row, 9, datetime.datetime.now())
                    self.driver.refresh()

        except (NoSuchElementException, TimeoutException) as e:
            print(e)
