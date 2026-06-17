import os
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import load_workbook


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"


# ------------------------
# Append row immediately
# ------------------------
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
            mode="a",
            engine="openpyxl",
            if_sheet_exists="overlay"
        ) as writer:

            ws = writer.book.active

            start_row = ws.max_row

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=start_row
            )


# ------------------------
# Extract one location
# ------------------------
def scrape_location(driver, url):

    phone = ""
    address = ""
    operating_hours = ""

    try:

        driver.get(url)

        wait = WebDriverWait(driver, 20)

        wait.until(
            EC.presence_of_element_located(
                (By.TAG_NAME, "body")
            )
        )

        time.sleep(5)

        # -------------------
        # ADDRESS
        # -------------------
        try:

            addr = driver.find_element(
                By.TAG_NAME,
                "address"
            )

            lines = addr.find_elements(
                By.TAG_NAME,
                "p"
            )

            address = ", ".join(
                [
                    x.text.strip()
                    for x in lines
                    if x.text.strip()
                ]
            )

        except:
            pass


        # -------------------
        # PHONE
        # -------------------
        try:

            phone = driver.find_element(
                By.XPATH,
                '//a[contains(@href,"tel:")]'
            ).text.strip()

        except:
            pass


        # -------------------
        # OPERATING HOURS
        # -------------------
        hours = []

        try:

            hours_section = driver.find_element(
                By.XPATH,
                '//h3[text()="Hours"]/parent::*'
            )

            rows = hours_section.find_elements(
                By.XPATH,
                './/div[contains(@class,"70qvj9")]'
            )

            for row in rows:

                vals = row.find_elements(
                    By.TAG_NAME,
                    "p"
                )

                if len(vals) >= 2:

                    day = vals[0].text.strip()
                    hour = vals[1].text.strip()

                    hours.append(
                        f"{day}: {hour}"
                    )

            operating_hours = " | ".join(hours)

        except:
            pass


    except Exception as e:

        print("Error:", e)

    return phone, address, operating_hours


# ------------------------
# Main
# ------------------------
input_df = pd.read_excel(INPUT_FILE)

driver = webdriver.Chrome(
    service=Service(
        ChromeDriverManager().install()
    )
)

try:

    for index, row in input_df.iterrows():

        featureid = row["featureid"]
        url = row["website"]

        print(f"\nProcessing {featureid}")

        phone, address, hours = scrape_location(
            driver,
            url
        )

        result = {
            "featureid": featureid,
            "website": url,
            "Phone": phone,
            "Address": address,
            "Operating_Hours": hours
        }

        append_excel(result)

        print("Saved")


finally:

    driver.quit()


print("\nCompleted")