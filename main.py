# For each row in the Excel file, fill the form with appropriate values to 
# First Name, Last Name, Sales Target and Sales


# Imports
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random

class ExcelHandling:
    
    def getDF(excelDataset):
        df = pd.read_excel(excelDataset)
        return df
    
class BrowserHandling:

    def getDriver(webLink):
        # setting up WebDriver to interact with given website in Edge browser
        driver = webdriver.Edge(service=Service(r"C:\Users\LukeFinch\Downloads\edgedriver_win64\msedgedriver.exe"))
        driver.get(webLink)

        return driver

    def logIn(df, driver):
        # Username
        driver.find_element(By.NAME, "username").send_keys("maria")

        # Password
        driver.find_element(By.NAME, "password").send_keys("thoushallnotpass")

        # Click Log In
        driver.find_element(By.XPATH, "//button[contains(text(),'Log in')]").click()
    
    def fillOutForm(df, driver):
        
        ### Now we can fill out the form for each row in the Excel sheet
        for index, row in df.iterrows():
            # Creating variables for each entry in a row
            fn = row['First Name']
            ln = row['Last Name']
            target = str(row['Sales Target'])
            result = row['Sales Result']

            # Typing info into form
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "firstname")))
            driver.find_element(By.NAME, "firstname").send_keys(fn)
            driver.find_element(By.NAME, "lastname").send_keys(ln)
            driver.find_element(By.NAME, "salesresult").send_keys(result)
            
            # Selecting from dropdown menu for Sales Target entry
            select = Select(driver.find_element(By.ID, "salestarget"))
            select.select_by_value(target)

            # Submitting form
            driver.find_element(By.XPATH, "//button[contains(text(),'Submit')]").click()

        driver.quit()

    def orderRobot(df, driver):
        # Switching over to the "Order your robot!" page
        driver.find_element(By.XPATH, "//a[@href='#/robot-order']").click()

        # Need to accept terms and conditions


        # Selecting a robot head
        select = Select(driver.find_element(By.NAME, "head"))
        select.select_by_value(str(random.randint(1, 6)))

        # Clicking a robot body (multiple choice)
        driver.find_element(By.ID, "id-body-" + str(random.randint(1, 6))).click() # Randomly choosing between 6 multiple choice options

        # Inputting robot legs
        driver.find_element(By.NAME, "legs").send_keys(str(random.randint(1, 4))) # Randomly picking the amount of legs through 1-4

        # Inputting shipment address
        driver.find_element(By.NAME, "address").send_keys("1324 Trapp Ln.")

        # Previewing order
        driver.find_element(By.XPATH, "//button[contains(text(),'Preview')]").click()

        # Submitting order
        driver.find_element(By.XPATH, "//button[contains(text(),'Order')]").click()

        driver.quit()


def main():
   df = ExcelHandling.getDF("SalesData.xlsx")
   driver = BrowserHandling.getDriver("https://robotsparebinindustries.com/")

   BrowserHandling.logIn(df, driver)
   BrowserHandling.fillOutForm(df, driver)
   # BrowserHandling.orderRobot(df, driver) ### EXTRA



if __name__ == "__main__":
    main()
