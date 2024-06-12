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

class ExcelHandling:
    
    def getDF(excelDataset):
        df = pd.read_excel(excelDataset)
        return df
    
class BrowserHandling:

    def driverSetUp(webLink):
        # setting up WebDriver to interact with given website in Edge browser
        driver = webdriver.Edge(service=Service(r"C:\Users\LukeFinch\Downloads\edgedriver_win64\msedgedriver.exe"))
        driver.get(webLink)

        return driver

    def fillOutForm(df, driver):
        ### Have to login first:

        # Username
        driver.find_element(By.NAME, "username").send_keys("maria")

        # Password
        driver.find_element(By.NAME, "password").send_keys("thoushallnotpass")

        # Click Log In
        driver.find_element(By.XPATH, "//button[contains(text(),'Log in')]").click()
        

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
        


def main():
   df = ExcelHandling.getDF("SalesData.xlsx")
   driver = BrowserHandling.driverSetUp("https://robotsparebinindustries.com/")

   BrowserHandling.fillOutForm(df, driver)



if __name__ == "__main__":
    main()
