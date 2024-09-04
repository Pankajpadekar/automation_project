from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time


username = "Chintan.K"
password = "Solutions@999"

url = "https://eintel.xpient.com/ffco/webcore/"

# service = Service(executable_path=r"E:\Projects\Automation_project\automation_project_selenium\geckodriver.exe")
driver = webdriver.Firefox()
driver.get(url)
driver.find_element(By.ID, "dnn_ctr362_Login_Login_DNN_txtUsername").send_keys(username)
driver.find_element(By.ID, "dnn_ctr362_Login_Login_DNN_txtPassword").send_keys(password)
driver.find_element(By.ID, "dnn_ctr362_Login_Login_DNN_cmdLogin").click()

# login Successfull


iframe = WebDriverWait(driver, 10).until(
    EC.frame_to_be_available_and_switch_to_it((By.ID, "rptFrame"))
)


dropdown = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ddlReports"))
)

select = Select(dropdown)


select.select_by_visible_text("SALES: Menu Mix By Major Minor Detail")

start_date = input("Enter the start date (mm/dd/yy): ")
end_date = input("Enter the end date (mm/dd/yy): ")


start_date_field = driver.find_element(By.ID, "RV1_ctl04_ctl03_txtValue")
driver.execute_script("arguments[0].value = arguments[1];", start_date_field, start_date)


end_date_field = driver.find_element(By.ID, "RV1_ctl04_ctl05_txtValue")
driver.execute_script("arguments[0].value = arguments[1];", end_date_field, end_date)


export_button = driver.find_element(By.XPATH, "//img[@id='RV1_ctl05_ctl04_ctl00_ButtonImg']")
export_button.click()

excel_link = driver.find_element(By.CSS_SELECTOR, "a[title='Excel'][alt='Excel']")
excel_link.click()


time.sleep(20)
