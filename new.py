import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains




# Create a new workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Data Sheet"
headerRow = ["Manufacturer", "Model", "Voltage", "Note", "Cover", "Lead Position",
              "Cover Requirements", "Components", "kWh Capacity", "Plate Rate (Ah)",
              "Ampere-Hour Capacity", "Weight", "Battery Chemistry", "Tray", "Dimensions"]
sheet.append(headerRow)
# set up the headless driver
driver = webdriver.Chrome()
driver.get("http://www.dekalifttruckguide.com/")
time.sleep(5)
wait = WebDriverWait(driver, 10)

action = ActionChains(driver)
manufacturerButton = driver.find_element(By.ID, "combo-1014-trigger-picker")
modelButton = driver.find_element(By.ID, "combo-1015-trigger-picker")
voltageButton = driver.find_element(By.ID, "combo-1016-trigger-picker")
compartmentSizeButton = driver.find_element(By.ID, "combo-1017-trigger-picker")
batteryFamilyButton = driver.find_element(By.ID, "combo-1018-trigger-picker")
manufacturerButton.click()
time.sleep(2)
manufacturer = driver.find_element(By.ID, 'combo-1014-picker-listEl')
manufacturerList = manufacturer.find_elements(By.XPATH,"./li")
manufacturerListLength = len(manufacturerList)
print("manufacturerListLength =", manufacturerListLength)
# loop through all of the manufacturers
for manufacturerIndex in range(1, manufacturerListLength + 1):
    if manufacturerIndex > 1:
        manufacturerButton.click()
        time.sleep(2)
    manufacturer = driver.find_element(By.ID, 'combo-1014-picker-listEl')
    manufacturerList = manufacturer.find_elements(By.XPATH,"./li")
    manufacturerList[manufacturerIndex].click()
    time.sleep(2)
    modelButton.click()
    time.sleep(2)
    model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
    modelList = model.find_elements(By.XPATH, './li')
    modelListLength = len(modelList)
    print("modelListLength =", modelListLength)
    # loop through all of the models
    for modelIndex in range(modelListLength):
        if modelListLength > 1:
            modelButton.click()
            time.sleep(2)
            model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
            modelList = model.find_elements(By.XPATH, './li')
            modelList[modelIndex].click()
            time.sleep(2)
        voltageButton.click()
        time.sleep(2)
        voltage = driver.find_element(By.ID, "combo-1016-picker-listEl")
        voltageList = voltage.find_elements(By.XPATH, "./li")
        voltageListLength = len(voltageList)
        voltageButton.click()
        time.sleep(2)
        print("voltageListLength =", voltageListLength)
        # loop through all of the voltages
        for voltageIndex in range(voltageListLength):
            if voltageListLength > 1:
                voltageButton.click()
                time.sleep(2)
                voltage = driver.find_element(By.ID, "combo-1016-picker-listEl")
                voltageList = voltage.find_elements(By.XPATH, "./li")
                voltageList[voltageIndex].click()
                time.sleep(2)
            compartmentSizeButton.click()
            time.sleep(2)
            ## the compartmentSizeList printed 1 when it should be 3
########################################################
            

            items = driver.find_elements(By.XPATH,"//ul[@id='combo-1017-picker-listEl']/li")
            time.sleep(2)
            # loop through all of the compartment sizes
            for item in items:
                #if len(items) > 1:
                action.move_to_element(item).perform()
                time.sleep(2)
                item.click()

                time.sleep(2)
                batteryFamilyButton.click()
                time.sleep(2)
                batteryFamily = driver.find_element(By.ID, "combo-1018-picker-listEl")
                batteryFamilyList = batteryFamily.find_elements(By.XPATH, "./li")
                batteryFamilyListLength = len(batteryFamilyList)
                batteryFamilyButton.click()
                time.sleep(2)
                print("batteryFamilyListLength =", batteryFamilyListLength)
                # loop through all of the battery families
                # for batteryFamilyIndex in range(batteryFamilyListLength):
                    # if batteryFamilyListLength > 1:
                batteryFamilyButton.click()
                time.sleep(2)
                batteryFamily = driver.find_element(By.ID, "combo-1018-picker-listEl")
                batteryFamilyList = batteryFamily.find_elements(By.XPATH, "./li")
                action.move_to_element(batteryFamilyList[0]).perform()
                time.sleep(2)
                batteryFamilyList[0].click()
                time.sleep(2)
                manufacturerText = driver.find_element("id", "combo-1014-inputEl").get_attribute("value")
                modelText = driver.find_element("id","combo-1015-inputEl").get_attribute("value")
                voltageText = driver.find_element("id", "combo-1016-inputEl").get_attribute("value")
                noteText = driver.find_element("id","textarea-1024-inputEl").get_attribute("value")
                coverText = driver.find_element("id","textfield-1026-inputEl").get_attribute("value")
                leadPositionText = driver.find_element("id","textfield-1027-inputEl").get_attribute("value")
                coverRequirementsText =driver.find_element("id","textfield-1029-inputEl").get_attribute("value")
                componentsText = driver.find_element("id", "combo-1017-inputEl").get_attribute("value")
                kWhCapacityText = driver.find_element("id", "textfield-1033-inputEl").get_attribute("value")
                plateRateAhText = driver.find_element("id","textfield-1034-inputEl" ).get_attribute("value")
                ampereHourCapacityText = driver.find_element("id","textfield-1035-inputEl").get_attribute("value")
                weightText = driver.find_element("id", "textfield-1036-inputEl").get_attribute("value")
                batteryChemistryText = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                trayText = driver.find_element("id", "textfield-1040-inputEl").get_attribute("value")
                dimensionsText = driver.find_element("id","textfield-1041-inputEl").get_attribute("value")
                # Write the values to the sheet
                dataRow = [manufacturerText, modelText, voltageText, noteText, coverText, leadPositionText, coverRequirementsText, componentsText, kWhCapacityText, plateRateAhText, ampereHourCapacityText, weightText, batteryChemistryText, trayText, dimensionsText] 
                sheet.append(dataRow)
                workbook.save("example.xlsx")  
                time.sleep(5)                      
driver.quit()
