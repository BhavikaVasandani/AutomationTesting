import os,unittest,time,xlrd
from selenium import webdriver
my_path=os.path.abspath(os.path.dirname(__file__))
path=os.path.join(my_path,"../Drivers/chromedriver.exe")
screenShot=os.path.join(my_path,'../ScreenPrint/')
import HtmlTestRunner
class ChromeLaunch(unittest.TestCase):
    @classmethod
    def setUpClass(cls):

        cls.driver=webdriver.Chrome(path)
        driver=cls.driver
        time.sleep(2)
        driver.maximize_window()


    def test_correct_detail(self):
        book = xlrd.open_workbook("C:/Users/amart/PycharmProjects/untitled1/Programs/data.xls")

        first_sheet =book.sheet_by_index(0)
        name = str(first_sheet.cell(1, 0).value)
        gender = first_sheet.cell(1, 1).value
        education = first_sheet.cell(1, 2).value
        sem = int(first_sheet.cell(1, 3).value)
        message = 1
        self.driver.get(
            "https://docs.google.com/forms/u/0/d/e/1FAIpQLSfBSJRUQeDQWmQ3XRbGFIIuhJqBzLEuyWdQvz6ha9igCcf3bQ/formResponse")
        time.sleep(30)
        #Please enter your email and password
        self.driver.find_element_by_xpath('//input[@type="email"]').send_keys("")
        time.sleep(30)
        self.driver.find_element_by_xpath('//*[@id="identifierNext"]').click()
        time.sleep(20)
        self.driver.find_element_by_xpath('//input[@type="password"]').send_keys("")
        time.sleep(30)
        self.driver.find_element_by_xpath('//*[@id="passwordNext"]').click()
        time.sleep(30)

        if name.isalpha():
            self.driver.find_element_by_name('entry.1353312019').send_keys(name)
        else:
            print("Invalid name")
        if gender == "Male" or gender == "M":
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[1]/label/div/div[1]/div/div[3]').click()
        elif gender=="Female" or gender=="F":
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[2]/label/div/div[1]/div/div[3]/div').click()
        else:
            print("Enter valid answer either Male, Female")
        if education=="Under Graduate" or "UG":
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[1]/div/label/div/div[1]/div[2]').click()
        elif education=="Post Graduate" or "PG":
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[2]/div/label/div/div[1]/div[2]').click()
        else:
            print("Enter valid answer either UG or PG")
        if sem == 1:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[1]/label/div/div[1]').click()
        elif sem==2:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[2]/label/div/div[1]').click()
        elif sem==3:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[3]/label/div/div[1]').click()
        elif sem==4:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[4]/label/div/div[1]').click()
        elif sem==5:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[5]/label/div/div[1]').click()
        elif sem==6:
            self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[6]/label/div/div[1]').click()
        else:
            print("Enter the sem in number from 1 - 6 only")
        self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[3]/span/span[2]').click()

        time.sleep(50)
        self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[3]/div[1]/div/div/span/span').click()
        time.sleep(30)


    def test_incorrect_Details(self):

        book = xlrd.open_workbook("C:/Users/amart/PycharmProjects/untitled1/Programs/data.xls")

        first_sheet = book.sheet_by_index(0)
        name = str(first_sheet.cell(2, 0).value)
        gender = first_sheet.cell(2, 1).value
        education = first_sheet.cell(2, 2).value
        sem = int(first_sheet.cell(2, 3).value)
        self.driver.get(
            "https://docs.google.com/forms/u/0/d/e/1FAIpQLSfBSJRUQeDQWmQ3XRbGFIIuhJqBzLEuyWdQvz6ha9igCcf3bQ/formResponse")
        time.sleep(30)



        if name.isalpha():
            self.driver.find_element_by_name('entry.1353312019').send_keys(name)
        else:
            print("Invalid name")

        if gender == "Male" or gender == "M":
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[1]/label/div/div[1]/div/div[3]').click()
        elif gender == "Female" or gender == "F":
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[2]/div/div[2]/div/span/div/div[2]/label/div/div[1]/div/div[3]/div').click()
        else:
            print("Enter valid answer either Male, Female")
        if education == "Under Graduate" or "UG":
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[1]/div/label/div/div[1]/div[2]').click()
        elif education == "Post Graduate" or "PG":
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[2]/div/label/div/div[1]/div[2]').click()
        else:
            print("Enter valid answer either UG or PG")
        if sem == 1:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[1]/label/div/div[1]').click()
        elif sem == 2:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[2]/label/div/div[1]').click()
        elif sem == 3:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[3]/label/div/div[1]').click()
        elif sem == 4:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[4]/label/div/div[1]').click()
        elif sem == 5:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[5]/label/div/div[1]').click()
        elif sem == 6:
            self.driver.find_element_by_xpath(
                '//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div/span/div/div[6]/label/div/div[1]').click()
        else:
            print("Enter the sem in number from 1 - 6 only")
        self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[3]/span/span[2]').click()

        time.sleep(50)
        self.driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[3]/div[1]/div/div/span/span').click()
        time.sleep(10)
    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
if __name__ == '__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner("C:/Users/amart/PycharmProjects/untitled1/Programs/Reports"))
