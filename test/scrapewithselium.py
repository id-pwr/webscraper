###### this is a bit of pain to set up
## here are the pip is
## pip install selenium
## pip install openpyxl
## to get selenium to work you need to download and add it to path.
### the issue I am having is on line 87 in the try: it doesnt want to click on the volts or the compartment size
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import openpyxl
#### it had problems with the compartment size I need to change that 

# Create a new workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Data Sheet"
header_row = ["Manufacturer", "Model", "Voltage", "Note", "Cover", "Lead Position",
              "Cover Requirements", "Components", "kWh Capacity", "Plate Rate (Ah)",
              "Ampere-Hour Capacity", "Weight", "Battery Chemistry", "Tray", "Dimensions"]
sheet.append(header_row)
# set up the headles driver
driver = webdriver.Chrome()
driver.get('http://www.dekalifttruckguide.com/')
time.sleep(2)
manufacture_btn = driver.find_element('id','combo-1014-trigger-picker')
manufacture_btn.click()
time.sleep(2)
manufacturer_ul = driver.find_element(By.ID, 'combo-1014-picker-listEl')
manufacture_li = manufacturer_ul.find_elements(By.XPATH,"./li")
x = len(manufacture_li)
manufacture_li[1].click()
## x = 38 
## loop throught all of the manufavures 
for index_one in range(x):
    if index_one == 0:
        time.sleep(2)
        model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
        model_li = model.find_elements(By.XPATH, './li')
        y = len(model_li)
        time.sleep(2)
        ###loop through  all of the models
        #  I need to reset this when the loop is done
        for index in range(y):
            if index == 0:
                model_li[index].click()
                time.sleep(2)
                batt_fam_ul = driver.find_element(By.ID, "combo-1018-picker-listEl")
                time.sleep(2)
                batt_fam_il = batt_fam_ul.find_elements(By.XPATH, "./li")
                batt_fam_il[0].click()
                manufacturer_text = driver.find_element("id", "combo-1014-inputEl").get_attribute("value")
                model_text = driver.find_element("id","combo-1015-inputEl").get_attribute("value")
                voltage_text = driver.find_element("id", "combo-1016-inputEl").get_attribute("value")
                note_txt = driver.find_element("id","textarea-1024-inputEl").get_attribute("value")
                cover_txt = driver.find_element("id","textfield-1026-inputEl").get_attribute("value")
                lead_pos_txt = driver.find_element("id","textfield-1027-inputEl").get_attribute("value")
                cov_req_txt =driver.find_element("id","textfield-1029-inputEl").get_attribute("value")
                comp_txt = driver.find_element("id", "combo-1017-inputEl").get_attribute("value")
                kwh_txt = driver.find_element("id", "textfield-1033-inputEl").get_attribute("value")
                plate_rate_ah_txt = driver.find_element("id","textfield-1034-inputEl" ).get_attribute("value")
                ah_cap_txt = driver.find_element("id","textfield-1035-inputEl").get_attribute("value")
                s_wight_txt = driver.find_element("id", "textfield-1036-inputEl").get_attribute("value")
                batt_type_txt = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                tray_txt = driver.find_element("id", "textfield-1040-inputEl").get_attribute("value")
                dimensions_txt = driver.find_element("id","textfield-1041-inputEl").get_attribute("value")
                # Write the values to the sheet
                data_row = [manufacturer_text, model_text, voltage_text, note_txt, cover_txt, lead_pos_txt, cov_req_txt, comp_txt, kwh_txt, plate_rate_ah_txt, ah_cap_txt, s_wight_txt, batt_type_txt, tray_txt, dimensions_txt] 
                sheet.append(data_row)
                workbook.save("example.xlsx")
                time.sleep(4)
            else:
                # manufacture_btn = driver.find_element('id','combo-1014-trigger-picker')
                manufacture_btn.click()
                time.sleep(2)
                manufacturer_ul = driver.find_element(By.ID, 'combo-1014-picker-listEl')
                manufacture_li = manufacturer_ul.find_elements(By.XPATH,"./li")
                manufacture_li[index_one + 1].click()
                time.sleep(2)
                model_btn = driver.find_element(By.ID, "combo-1015-trigger-picker")
                model_btn.click()
                model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
                model_li = model.find_elements(By.XPATH, './li')
                model_li[index].click()
                time.sleep(2)
                voltage_btn = driver.find_element(By.ID, "combo-1016-trigger-picker")
                try:
                    voltage_ul = driver.find_element(By.ID, "combo-1016-picker-listEl")
                    voltage_il = voltage_ul.find_elements(By.XPATH, "./li")
                    voltage_il_length = len(voltage_il)
                    if voltage_il_length != 0:
                        for volts in range(voltage_il_length):
                            voltage_il[volts].click()
                            compartment_size_ul = driver.find_element(By.ID, "combo-1017-picker-listEl")
                            compartment_size_il = compartment_size_ul.find_elements(By.XPATH, "./li")
                            compartment_size_il_length = len(compartment_size_il)
                            if compartment_size_il_length != 0:
                                for comp in range(compartment_size_il_length):
                                    time.sleep(2)
                                    compartment_size_il[comp].click()
                except NoSuchElementException:
                    print("not found")
                else:
                    try:
                        comparment_size_ul = driver.find_element("id", "combo-1017-picker-listEl")
                        compartment_size_il = compartment_size_ul.find_elements(By.XPATH, "./li")
                        if comparment_size_il_lenght != 0:
                                # compartmant_size = driver.find_element(By.ID, "combo-1017-trigger-picker")
                                for comp in range(comparment_size_il_lenght):
                                    comparment_size_il[comp].click()
                    except: NoSuchElementException
                batt_fam_ul = driver.find_element(By.ID, "combo-1018-picker-listEl")
                time.sleep(2)
                batt_fam_il = batt_fam_ul.find_elements(By.XPATH, "./li")
                time.sleep(2)
                batt_fam_il[0].click()
                time.sleep(2)
                ### I need to grab all the info and save it to a file
                manufacturer_text = driver.find_element("id", "combo-1014-inputEl").get_attribute("value")
                model_text = driver.find_element("id","combo-1015-inputEl").get_attribute("value")
                voltage_text = driver.find_element("id", "combo-1016-inputEl").get_attribute("value")
                note_txt = driver.find_element("id","textarea-1024-inputEl").get_attribute("value")
                cover_txt = driver.find_element("id","textfield-1026-inputEl").get_attribute("value")
                lead_pos_txt = driver.find_element("id","textfield-1027-inputEl").get_attribute("value")
                cov_req_txt =driver.find_element("id","textfield-1029-inputEl").get_attribute("value")
                comp_txt = driver.find_element("id", "combo-1017-inputEl").get_attribute("value")
                kwh_txt = driver.find_element("id", "textfield-1033-inputEl").get_attribute("value")
                plate_rate_ah_txt = driver.find_element("id","textfield-1034-inputEl" ).get_attribute("value")
                ah_cap_txt = driver.find_element("id","textfield-1035-inputEl").get_attribute("value")
                s_wight_txt = driver.find_element("id", "textfield-1036-inputEl").get_attribute("value")
                batt_type_txt = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                tray_txt = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                dimensions_txt = driver.find_element("id","textfield-1041-inputEl").get_attribute("value")
                # Write the values to the sheet
                data_row = [manufacturer_text, model_text, voltage_text, note_txt, cover_txt, lead_pos_txt, cov_req_txt, comp_txt, kwh_txt, plate_rate_ah_txt, ah_cap_txt, s_wight_txt, batt_type_txt, tray_txt, dimensions_txt] 
                sheet.append(data_row)
                workbook.save("example.xlsx")
               
    else:
        manufacture_btn.click()
        time.sleep(2)
        manufacturer_ul = driver.find_element(By.ID, 'combo-1014-picker-listEl')
        manufacture_li = manufacturer_ul.find_elements(By.XPATH,"./li")      
        manufacture_li[index_one + 1].click()
        model_btn.click()
        model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
        model_li = model.find_elements(By.XPATH, './li')
        y = len(model_li)
        time.sleep(2)
        ###loop through  all of the models
        #  I need to reset this when the loop is done
        for index in range(y):
                time.sleep(2)
                model_btn = driver.find_element(By.ID, "combo-1015-trigger-picker")
                model_btn.click()
                time.sleep(2)
                model = driver.find_element(By.ID, 'combo-1015-picker-listEl')
                model_li = model.find_elements(By.XPATH, './li')
                model_li[index].click()
                time.sleep(2)
                voltage_btn = driver.find_element(By.ID, "combo-1016-trigger-picker")
                try:
                    voltage_ul = driver.find_element(By.ID, "combo-1016-picker-listEl")
                    voltage_il = voltage_ul.find_elements(By.XPATH, "./li")
                    voltage_il_length = len(voltage_il)
                    if voltage_il_length == 0:
                        print("do nothing")
                    else:
                        for volts in range(voltage_il_length):
                            voltage_il[volts].click()
                            comparment_size_ul = driver.find_element(By.ID, "combo-1017-picker-listEl")
                            comparment_size_il = comparment_size_ul.find_elements(By.XPATH, "./li")
                            comparment_size_il_lenght = len(comparment_size_il)
                            if comparment_size_il_lenght == 0:
                                print("do Nothing for comp")
                            else:
                                for comp in range(comparment_size_il_lenght):
                                    comparment_size_il[comp].click()
                except NoSuchElementException:
                    print("not found")
                else:
                    try:
                        comparment_size_ul = driver.find_element("id", "combo-1017-picker-listEl")
                        comparment_size_il = comparment_size_ul.find_elements(By.XPATH, "./li")
                        comparment_size_il_lenght = len(comparment_size_il)
                        if comparment_size_il_lenght != 0:
                                # compartmant_size = driver.find_element(By.ID, "combo-1017-trigger-picker")
                                for comp in range(comparment_size_il_lenght):
                                    comparment_size_il[comp].click()
                    except: NoSuchElementException
                time.sleep(2)
                batt_fam_ul = driver.find_element(By.ID, "combo-1018-picker-listEl")
                time.sleep(2)
                batt_fam_il = batt_fam_ul.find_elements(By.XPATH, "./li")
                time.sleep(2)
                batt_fam_il[0].click()
                time.sleep(2)
                ### I need to grab all the info and save it to a file
                manufacturer_text = driver.find_element("id", "combo-1014-inputEl").get_attribute("value")
                model_text = driver.find_element("id","combo-1015-inputEl").get_attribute("value")
                voltage_text = driver.find_element("id", "combo-1016-inputEl").get_attribute("value")
                note_txt = driver.find_element("id","textarea-1024-inputEl").get_attribute("value")
                cover_txt = driver.find_element("id","textfield-1026-inputEl").get_attribute("value")
                lead_pos_txt = driver.find_element("id","textfield-1027-inputEl").get_attribute("value")
                cov_req_txt =driver.find_element("id","textfield-1029-inputEl").get_attribute("value")
                comp_txt = driver.find_element("id", "combo-1017-inputEl").get_attribute("value")
                kwh_txt = driver.find_element("id", "textfield-1033-inputEl").get_attribute("value")
                plate_rate_ah_txt = driver.find_element("id","textfield-1034-inputEl" ).get_attribute("value")
                ah_cap_txt = driver.find_element("id","textfield-1035-inputEl").get_attribute("value")
                s_wight_txt = driver.find_element("id", "textfield-1036-inputEl").get_attribute("value")
                batt_type_txt = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                tray_txt = driver.find_element("id", "textfield-1039-inputEl").get_attribute("value")
                dimensions_txt = driver.find_element("id","textfield-1041-inputEl").get_attribute("value")
                # Write the values to the sheet
                data_row = [manufacturer_text, model_text, voltage_text, note_txt, cover_txt, lead_pos_txt, cov_req_txt, comp_txt, kwh_txt, plate_rate_ah_txt, ah_cap_txt, s_wight_txt, batt_type_txt, tray_txt, dimensions_txt] 
                sheet.append(data_row)
                workbook.save("example.xlsx")
                
            
driver.quit()


    