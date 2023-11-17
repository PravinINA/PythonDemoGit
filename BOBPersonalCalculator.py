from selenium import webdriver
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

import XLUtility

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.implicitly_wait(10)
driver.get("https://www.bankofbaroda.in/calculators/personal-loan-emi-calculator")
driver.find_element(By.XPATH,"//div[@class='privacy-warning acceptonclose']/div[@class='close']").click()
inptFile="D:\Automation\PyAutomation\InputFiles.xlsx"
optFile="D:\Automation\PyAutomation\OptFiles.xlsx"

row=XLUtility.getrowCount(inptFile,"PLCalcData")
print(row)

for r in range(3,row+1):
    prcAmount=XLUtility.readData(inptFile,"PLCalcData",r,1)
    rateOfInterest=XLUtility.readData(inptFile,"PLCalcData",r,2)
    tenure=XLUtility.readData(inptFile,"PLCalcData",r,3)
    print(prcAmount)
    act = ActionChains(driver);

    act.double_click(driver.find_element(By.XPATH,"//small[contains(text(),'₹')]")).perform()
    act.send_keys(prcAmount).perform()

    act.double_click(driver.find_element(By.XPATH,"(//em[@contenteditable='true'])[2]")).perform()
    act.send_keys(Keys.BACKSPACE).perform()
    act.send_keys(rateOfInterest).perform()

    act.double_click(driver.find_element(By.XPATH, "(//em[@contenteditable='true'])[3]")).perform()
    act.send_keys(tenure).perform()