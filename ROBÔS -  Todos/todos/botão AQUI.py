import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Initialize the webdriver
driver = webdriver.Chrome()

# Open the GUIA FINAL.py robot file
driver.get("file:///path/to/GUIA_FINAL.py")

# Loop through all iframes searching for the "AQUI" button
found = False
for frame in driver.find_elements(By.TAG_NAME, "iframe"):
    driver.switch_to.frame(frame)
    try:
        # Wait for the "AQUI" button to be clickable within this frame
        button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='pdfjs_internal_id_8R']")))
        button.click()
        found = True  # Button found, exit the loop
        break
    except NoSuchElementException:
        driver.switch_to.default_content()  # Switch back to parent frame

# If not found in any frame, handle the case
if not found:
    print("AQUI button not found in any iframe!")

# Close the PDF viewer iframe (if necessary)
driver.switch_to.default_content()

# Close the browser window
driver.quit()
