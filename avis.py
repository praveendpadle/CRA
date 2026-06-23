import os
import time
import pandas as pd
import undetected_chromedriver as uc

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"


# -----------------------------------
# Scroll (lazy loading)
# -----------------------------------
def scroll_page(driver):

    last_height = 0

    while True:

        new_height = driver.execute_script(
            "return document.body.scrollHeight"
        )

        if new_height == last_height:
            break

        driver.execute_script(
            "window.scrollTo(0, arguments[0])",
            new_height
        )

        time.sleep(2)

        last_height = new_height

    driver.execute_script(
        "window.scrollTo(0,0)"
    )

    time.sleep(1)


# -----------------------------------
# Append output
# -----------------------------------
def append_excel(row):

    df = pd.DataFrame([row])

    if not os.path.exists(OUTPUT_FILE):

        df.to_excel(
            OUTPUT_FILE,
            index=False
        )

    else:

        with pd.ExcelWriter(
            OUTPUT_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="overlay"
        ) as writer:

            start = writer.book.active.max_row

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=start
            )


# -----------------------------------
# Safe element text
# -----------------------------------
def safe_text(el):

    try:
        return el.text.strip()

    except:
        return ""


# -----------------------------------
# Scrape Budget
# -----------------------------------
def scrape_budget(driver, url):

    phone = ""
    address = ""
    hours = ""
    error = ""

    # try:

    driver.get(url)

    wait = WebDriverWait(driver, 30)

    wait.until(
        EC.presence_of_element_located(
            (By.TAG_NAME, "body")
        )
    )

    time.sleep(5)

    driver.execute_script(
        f"window.scrollTo(0,300)"
    )

        # --------------------
# ADDRESS
# --------------------
    time.sleep(2)  # give lazy loader time to populate
    driver.refresh()
    time.sleep(3)  # give lazy loader time to populate
    address = ""
    try:
    # Strategy 1: By data-testid
        element = driver.find_element(By.CSS_SELECTOR, "[data-testid='location-list-view-station-address-value']")
        address_text = element.text
        address = address_text
        print("Address:", address_text)
        if not address_text:
            element = driver.find_element(By.CSS_SELECTOR, "[data-testid='location-list-view-station-address-value']")
            address_text = element.text
            address = address_text
            print("Address:", address_text)
    except:
        print("Element not found with any selector")

    #Phone
    phone = ""
    
    try:
        element = driver.find_element(By.CSS_SELECTOR, "[data-testid='location-list-view-station-phone-values']")
        phone_text = element.text
        phone = phone_text
        print("Phone:", phone_text)
        if not phone_text:
            element = driver.find_element(By.CSS_SELECTOR, "[data-testid='location-list-view-station-phone-values']")
            phone_text = element.text
            phone = phone_text
            print("Phone:", phone_text)
    except:
        print("Phone Element not found with any selector")

     #Operating Hours
    hours = ""

    try:
        days = driver.find_elements(By.CSS_SELECTOR, "[data-testid='location-list-view-station-hours-of-operation-days']")
        time_slots = driver.find_elements(By.CSS_SELECTOR, "[data-testid='location-list-view-station-hours-of-operation-time-slot']")

        # Pair them together
        operating_hours = {}
        for day, time_slot in zip(days, time_slots):
            day_text = day.text.strip()
            time_text = time_slot.text.strip()
            operating_hours[day_text] = time_text

# Print results
        for day, hours in operating_hours.items():
            print(f"{day}: {hours}")
            hours += f"{day}: {hours}|"

        if not operating_hours:
            days = driver.find_elements(By.CSS_SELECTOR, "[data-testid='location-list-view-station-hours-of-operation-days']")
            time_slots = driver.find_elements(By.CSS_SELECTOR, "[data-testid='location-list-view-station-hours-of-operation-time-slot']")

            # Pair them together
            operating_hours = {}
            for day, time_slot in zip(days, time_slots):
                day_text = day.text.strip()
                time_text = time_slot.text.strip()
                operating_hours[day_text] = time_text

            # Print results
            for day, hours in operating_hours.items():
                print(f"{day}: {hours}")
                hours += f"{day}: {hours}|"
    except:
        print("Hours Element not found with any selector")

    return phone, address, hours, error

# -----------------------------------
# Driver
# -----------------------------------
options = uc.ChromeOptions()

options.add_argument("--start-maximized")

driver = uc.Chrome(
    options=options,
    version_main=149,
    use_subprocess=True
)


# -----------------------------------
# Read input
# -----------------------------------
df = pd.read_excel(INPUT_FILE)


try:

    for _, row in df.iterrows():

        featureid = str(
            row["featureid"]
        ).strip()

        website = str(
            row["website"]
        ).strip()

        if not website:
            continue

        print(f"\nProcessing → {featureid}")

        phone, address, hours, error = (
            scrape_budget(
                driver,
                website
            )
        )

        output = {
            "featureid": featureid,
            "website": website,
            "Phone": phone,
            "Address": address,
            "Operating_Hours": hours,
            "Error": error
        }

        append_excel(output)

        print("Saved")
        time.sleep(2)


finally:
    pass
    #driver.quit()


print("\nCompleted")