from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import time
from xlsxwriter import Workbook
import pandas as pd
options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()),
                          options = options)


driver.get('https://esearch.delhigovt.nic.in/Complete_search.aspx')
element = Select(driver.find_element("id","ctl00_ContentPlaceHolder1_ddl_sro_s"))
element.select_by_visible_text('Central -Asaf Ali (SR III)')
# Locality
locality = Select(driver.find_element("id","ctl00_ContentPlaceHolder1_ddl_loc_s"))
locality_names = []
data = pd.DataFrame()
for opt in locality.options:
    locality_names.append(opt.text)
locality_names.pop(0)
locality_data = [x for x in locality_names if "*" not in x]
def get_values():
    dataset_header = []
    dataset_rows = []
    table_header = driver.find_elements("xpath", "//*[@id='ctl00_ContentPlaceHolder1_gv_search']/tbody/tr[1]/th")
    for i in table_header:
        dataset_header.append(i.text)
    table_row = driver.find_elements("xpath", "//*[@id='ctl00_ContentPlaceHolder1_gv_search']/tbody/tr")
    for i in range(2,len(table_row)+1):
        row_value = driver.find_elements("xpath", "//*[@id='ctl00_ContentPlaceHolder1_gv_search']/tbody/tr["+str(i)+"]/td")
        temp = []
        for j in row_value:
            temp.append(j.text)
        dataset_rows.append(temp)
    return dataset_header,dataset_rows[:-2]


for locality in locality_data[:1]:
    driver.get('https://esearch.delhigovt.nic.in/Complete_search.aspx')
    element = Select(driver.find_element("id","ctl00_ContentPlaceHolder1_ddl_sro_s"))
    element.select_by_visible_text('Central -Asaf Ali (SR III)')
    element = Select(driver.find_element("id","ctl00_ContentPlaceHolder1_ddl_loc_s"))
    element.select_by_visible_text(locality)
    element = Select(driver.find_element("id","ctl00_ContentPlaceHolder1_ddl_year_s"))
    element.select_by_visible_text("2021-2022")
    try:
        WebDriverWait(driver, 10000).until(EC.text_to_be_present_in_element((By.ID, "ctl00_ContentPlaceHolder1_Label1"), 'Search Result-'))
        first_regno = driver.find_element("xpath","//*[@id='ctl00_ContentPlaceHolder1_gv_search']/tbody/tr[2]/td[1]").text
        dataset_header,dataset_rows = get_values()
        data = pd.concat([data,pd.DataFrame(dataset_rows,columns=dataset_header)], ignore_index = True)
        tot_no_pages = driver.find_element("xpath", "//*[@id='ctl00_ContentPlaceHolder1_gv_search_ctl13_lblTotalNumberOfPages']")
        for pg_no in range(2,int(tot_no_pages.text)+1):
            driver.find_element("xpath","//*[@name='ctl00$ContentPlaceHolder1$gv_search$ctl13$txtGoToPage']").send_keys(Keys.BACK_SPACE,Keys.BACK_SPACE, pg_no,Keys.ENTER)
            WebDriverWait(driver, 10000).until_not(EC.text_to_be_present_in_element(("xpath", "//*[@id='ctl00_ContentPlaceHolder1_gv_search']/tbody/tr[2]/td[1]"), first_regno))
            dataset_header,dataset_rows = get_values()
            data = pd.concat([data,pd.DataFrame(dataset_rows,columns=dataset_header)], ignore_index = True)
    except NoSuchElementException:
        pass
writer = pd.ExcelWriter(r'C:\Users\anura\OneDrive\Desktop\Book1.xlsx', engine='xlsxwriter')
data.to_excel(writer, sheet_name='sheet', index=False)
writer.close()
driver.quit()