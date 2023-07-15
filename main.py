import uuid
import os
from dotenv import load_dotenv
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import pygsheets

load_dotenv()

def save_screenshot(driver, path):
    original_size = driver.get_window_size()
    required_width = driver.execute_script('return document.body.parentNode.scrollWidth')
    required_height = driver.execute_script('return document.body.parentNode.scrollHeight')
    driver.set_window_size(required_width, required_height)
    # driver.save_screenshot(path)  # has scrollbar
    driver.find_element(By.TAG_NAME, 'body').screenshot(path)  # avoids scrollbar
    driver.set_window_size(original_size['width'], original_size['height'])

# Read from Google sheet
gc = pygsheets.authorize(service_file = os.getenv("GOOGLE_SHEET_API_KEY_FILE"))

sht = gc.open_by_url(os.getenv("TEST_PLAN"))
wks_list = sht.worksheets()

target = {"uuid":"", "host":"", "case":"", "auth":"", "snapshot":"", "steps":[], "RWD":""}

## 環境變數
# UUID
target["uuid"] = uuid.uuid1()

# host
wks = sht[0] 
host = wks.cell("A2")
target["host"] = host.value

# case
cases = wks.cell("B2")
target["case"] = cases.value

# auth
auth = wks.cell("C2")
target["auth"] = auth.value

# snapshot
snapshot = wks.cell("D2")
target["snapshot"] = snapshot.value

# RWD
rwd = wks.cell("E2")
target["RWD"] = rwd.value

# plan
wks = sht[1] 
index_column = 0
list_row_data = wks.get_row(1)
for obj in list_row_data:
    index_column += 1
    if obj == target["case"]:
        plan = obj
        break

list_col = wks.get_col(index_column)

for obj in list_col[1:]:
    if (obj != ""):
        target["steps"].append(obj)

## Webdriver
options = webdriver.ChromeOptions()
options.add_argument("incognito")

driver = webdriver.Chrome(options = options)
driver.maximize_window()

url_target = target["host"]
if(bool(target["auth"])):
  url_target = url_target + "/tplanet_signin.html"
  driver.get(url_target)
  driver.find_element(By.ID, value = "email").send_keys(os.getenv("TESTER_EMAIL"))
  driver.find_element(By.ID, value = "password").send_keys(os.getenv("TESTER_PWD"))
  driver.find_element(By.TAG_NAME, "form").submit()
  time.sleep(5)

## Plan
for step in target["steps"]:
    driver.get(target["host"] + step)
    time.sleep(5)

    # Snapshot
    if (bool(target["snapshot"])):
      # Scroll
      driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

      # 擷取完整網頁截圖
      save_screenshot(driver, "output/" + str(target["uuid"]) + ".png")

## TODO Report

## TODO: Notify

driver.quit()
