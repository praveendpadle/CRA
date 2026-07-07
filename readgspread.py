import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# 1. Mount your authenticated Chrome profile
options = Options()
# Update this path to your actual Chrome User Data folder
user_data_dir = r"C:\Users\YourUser\AppData\Local\Google\Chrome\User Data" 
options.add_argument(f"user-data-dir={user_data_dir}")
options.add_argument("profile-directory=Default")

driver = webdriver.Chrome(options=options)

try:
    # 2. Open the standard edit URL of your private sheet
    sheet_id = "YOUR_PRIVATE_SHEET_ID"
    driver.get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    
    # Wait for the HTML5 canvas and scripts to fully render
    time.sleep(5) 
    
    # 3. Force focus onto the sheet canvas by clicking a fixed offset
    # This clicks an arbitrary point (x=200, y=300) below the toolbar
    ActionChains(driver).move_by_offset(200, 300).click().perform()
    time.sleep(1)
    
    # 4. Simulate Ctrl+A (Select All) then Ctrl+C (Copy)
    # Note: If running on macOS, change Keys.CONTROL to Keys.COMMAND
    actions = ActionChains(driver)
    
    # Select All
    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    time.sleep(1) # Brief pause to allow the UI to highlight cells
    
    # Copy to clipboard
    actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
    time.sleep(1) # Brief pause for the OS to register the clipboard data
    
    # 5. Read the system clipboard directly into a Pandas DataFrame
    df = pd.read_clipboard()
    
    print("Data successfully scraped to memory:")
    print(df.head())

finally:
    driver.quit()