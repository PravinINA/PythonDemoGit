import os

import openpyxl as openpyxl
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.implicitly_wait(10)
driver.get("https://www.bankofbaroda.in/")
driver.find_element(By.XPATH,"//div[@class='privacy-warning acceptonclose']/div[@class='close']").click()
act=ActionChains(driver);
driver.find_element(By.XPATH,"//a[@class='search-popup']").click()
driver.find_element(By.XPATH,"//input[@placeholder='Looking for something specific?']").send_keys("Home Loan")
search_result=[]


search_result_option=driver.find_elements(By.XPATH,"//ul[@class='search-result-list']/li/h4")
print(len(search_result_option))


for search_item in search_result_option:
    driver.execute_script("arguments[0].scrollIntoView(true);", search_item )
    print(search_item.text)
    search_result.append(search_item.text)

print(search_result)
#driver.get_screenshot_as_file(os.getcwd()+"\\homeloansearch.png")
cookies=driver.get_cookies();
for c in cookies:
    print(c)

#book=openpyxl.load_workbook("/search.xlsx")





