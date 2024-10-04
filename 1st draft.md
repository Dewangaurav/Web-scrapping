# Web-scrapping

from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import os

# Specify the path to geckodriver
geckodriver_path = r"C:\WebDriver\geckodriver.exe"

# Initialize Firefox options (without headless mode)
firefox_options = FirefoxOptions()

# (Optional) Add Firefox profile path (if you want to use a specific Firefox profile)
firefox_options.profile = r"C:\Users\dewagaur\AppData\Roaming\Mozilla\Firefox\Profiles\83nybmxd.default-esr"

# Set a custom download directory (if needed)
firefox_options.set_preference("browser.download.dir", r"C:\Users\dewagaur\Downloads")
firefox_options.set_preference("browser.download.folderList", 2)
firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")

# Initialize the WebDriver using Service and options
service = FirefoxService(executable_path=geckodriver_path)
driver = webdriver.Firefox(service=service, options=firefox_options)

# Open the webpage
url = "https://mars-admin.aka.amazon.com/batch-statistics"
driver.get(url)

# Increase the timeout for elements to load
wait = WebDriverWait(driver, 30)

# Add a delay to give the page some time to load completely
time.sleep(30)

# Debugging print to check if the page loaded
print("Page loaded")

# Step 1: Click the scroll menu to open dropdown (First click)
try:
    print("Attempting to click scroll menu")
    scroll_menu = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="adPrograms"]')))
    scroll_menu.click()
    print("Scroll menu clicked")

    print("Attempting to select Moderation option")
    moderation_option = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div/div/div/div/div/nav[1]/div/ul/li[1]/span/div/div/ul/li[3]/a")))
    moderation_option.click()
    print("Moderation option selected")
except Exception as e:
    print(f"Error in Step 1: {e}")

# Step 2: Click the 'Marketplace' scroll menu and select 'EN'
try:
    print("Attempting to click Marketplace scroll menu")
    marketplace_menu = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//li[//label[contains(text(),'Marketplace')]]/following-sibling::li//button[@id='language']")))
    marketplace_menu.click()
    print("Marketplace menu clicked")

    print("Attempting to select EN option")
    en_option = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div/div/div/div/div/nav[2]/div/ul[2]/li[2]/span/div/div/ul/li[2]/a")))
    driver.execute_script("arguments[0].scrollIntoView(true);", en_option)
    en_option.click()
    print("EN option selected")
except Exception as e:
    print(f"Error in Step 2: {e}")
    try:
        actions = ActionChains(driver)
        actions.move_to_element(en_option).click().perform()
        print("EN option clicked using ActionChains!")
    except Exception as e:
        print(f"Error: {e}")

# Step 3: Press the first button
try:
    first_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div/div/div/div/div/nav[1]/div/ul/li[2]/span/button[2]")))
    first_button.click()
    print("First button clicked")
except Exception as e:
    print(f"Failed to click first button: {e}")

# Step 4: Press the second button
try:
    second_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div/div/div/div/div/nav[1]/div/ul/li[2]/span/button[1]")))
    second_button.click()
    print("Second button clicked")
except Exception as e:
    print(f"Failed to click second button: {e}")

# Wait for data to load after button press
time.sleep(10)

# Initialize variables in case elements are not found
ad_queue, week, volumes, sla_breached = None, None, None, None

# 1. Extract "Ad Queue" (Extract all Ad Queues)
try:
    ad_queue_elements = wait.until(EC.presence_of_all_elements_located(  # Modified
        (By.XPATH, "//div[@class='container-fluid']//div[@class='title row']")))
    ad_queue = [element.text for element in ad_queue_elements]  # Modified
    print("Ad Queues:", ad_queue)
except Exception as e:
    print(f"Ad Queue not found: {e}")

# 2. Extract "Week" (Extract all Weeks)
try:
    week_elements = wait.until(EC.presence_of_all_elements_located(  # Modified
        (By.XPATH, "//div[contains(@class, 'container-fluid')]//div[@class='title row']/following-sibling::div[1]")))
    week = [element.text for element in week_elements]  # Modified
    print("Weeks:", week)
except Exception as e:
    print(f"Week not found: {e}")

# 3. Extract "Volumes" (Extract all Volumes)
try:
    volumes_elements = wait.until(EC.presence_of_all_elements_located(  # Modified
        (By.XPATH, "//div[contains(@class, 'container-fluid')]//div[contains(@class, 'row')]//div/div[3]/div[1]/span/span[1]")))
    volumes = [element.text for element in volumes_elements]  # Modified
    print("Volumes:", volumes)
except Exception as e:
    print(f"Volumes not found: {e}")

# 4. Extract "SLA breached" (Extract all SLA breaches)
try:
    sla_breached_elements = wait.until(EC.presence_of_all_elements_located(  # Modified
        (By.XPATH, "//div[contains(@class, 'container-fluid')]//div[contains(@class, 'row')]//div/div[2]/div[2]/div/div[2]")))
    sla_breached = [element.text for element in sla_breached_elements]  # Modified
    print("SLA Breached:", sla_breached)
except Exception as e:
    print(f"SLA breached not found: {e}")

# Close the WebDriver after scraping is complete
driver.quit()

# Store the extracted data in a pandas DataFrame
data = {
    "Ad Queue": ad_queue,
    "Week": week,
    "Volumes": volumes,
    "SLA breached": sla_breached
}

df = pd.DataFrame(data)

# Specify the path where you want to save the Excel file
excel_file_path = os.path.join(r"C:\Users\dewagaur\Downloads", "excel_output.xlsx")

# Write the data to an Excel file
df.to_excel(excel_file_path, index=False)

print(f"Data has been successfully written to {excel_file_path}")
