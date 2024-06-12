# For each row in the Excel file, fill the form with appropriate values to 
# First Name, Last Name, Sales Target and Sales


# Imports
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

def main():
    # Imported the excel file as a CSV file so I can iterate through it
    df = pd.read_excel('SalesData.xlsx')


    # setting up WebDriver to interact with given website
    driver = webdriver.Chrome(service=Service('path_to_chromedriver'))
    driver.get('https://robotsparebinindustries.com/#/')

    for row in df:
        # Creating variables for each entry in a row
        fn = row['First Name']
        ln = row['Last Name']
        target = '$' + row['Sales Target']
        result = row['Sales Result']

        # Typing info into form
        driver.find_element(By.NAME, "First Name").send_keys(fn)
        driver.find_element(By.NAME, "Last Name").send_keys(ln)
        driver.find_element(By.NAME, "Sales Result ($)").send_keys(result)
        
        # Selecting from dropdown menu for Sales Target entry
        select = Select(driver.find_element(By.NAME, "Sales Target ($)"))
        select.select_by_visible_text(target)

        # Submitting form
        driver.find_element(By.NAME, "SUBMIT").click()

    driver.quit()

if __name__ == "__main__":
    main()
