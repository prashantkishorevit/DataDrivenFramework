import unittest
import openpyxl as pyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from py._xmlgen import html


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(5)
        self.driver.get("http://newtours.demoaut.com/")
        self.read_excel()


    def login(self,username,password):
        self.driver.find_element(By.NAME, "userName").send_keys(username)
        self.driver.find_element(By.NAME,"password").send_keys(password)
        self.driver.find_element(By.NAME,"login").click()
        my_title = self.driver.title
        if my_title.startswith("Find a Flight"):
            self.driver.find_element(By.LINK_TEXT,"SIGN-OFF").click()
            return 'VALID_USER'
        else:
            self.assertTrue(my_title.startswith("Sign-on"))
            return "IN-VALID USER"


    def read_excel(self):
        file = "C:\Code_Python\Data_Driven_FrameWork\Data_Driven_1\Data\Book1.xlsx" # C:\Code_Python\Data_Driven_FrameWork\Data_Driven_1\Data\Book1.xlsx
        wb = pyxl.load_workbook(file) # C:\Code_Python\Data_Driven_FrameWork\Data_Driven_1\data_driven_test_1.py
        sheet1=wb['Sheet1']

        for c1,c2,c3 in sheet1['A2':'C4']:
            c3.value = self.login(c1.value,c2.value)


        wb.save(file)


if __name__ == '__main__':
    unittest.main()
