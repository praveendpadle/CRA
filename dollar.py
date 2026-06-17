import os
import time
import pandas as pd
import undetected_chromedriver as uc

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

# --------------------------------
# Append extracted row
# --------------------------------
def append_to_excel(row):

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

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=ws.max_row
            )


# --------------------------------
# Extract page
# --------------------------------
def scrape_store(driver, url):

    phone = ""
    address = ""
    operating_hours = ""

    try:

        driver.get(url)

        wait = WebDriverWait(driver, 30)

        wait.until(
            EC.presence_of_element_located(
                (By.TAG_NAME, "body")
            )
        )

        time.sleep(6)

        driver.execute_script(
            f"window.scrollTo(0,1000)"
        )
        time.sleep(2)  # give lazy loader time to populate
        # -------------------
        # ADDRESS
        # -------------------
        try:

            addr = driver.find_element(
                By.TAG_NAME,
                "address"
            )

            address = ", ".join(
                [
                    p.text.strip()
                    for p in addr.find_elements(
                        By.TAG_NAME,
                        "p"
                    )
                    if p.text.strip()
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
        # HOURS
        # -------------------
        try:

            hours = []

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


    except Exception as ex:

        print("Page Error:", ex)

    return phone, address, operating_hours


# --------------------------------
# Driver
# --------------------------------
options = uc.ChromeOptions()

options.add_argument("--start-maximized")

driver = uc.Chrome(
    options=options,
    use_subprocess=True
)


# --------------------------------
# Read input
# --------------------------------
df = pd.read_excel(INPUT_FILE)

try:

    for _, row in df.iterrows():

        featureid = row["featureid"]
        website = row["website"]

        print(f"\nProcessing -> {featureid}")

        phone, address, hours = scrape_store(
            driver,
            website
        )

        result = {
            "featureid": featureid,
            "website": website,
            "Phone": phone,
            "Address": address,
            "Operating_Hours": hours
        }

        append_to_excel(result)

        print("Saved")

finally:

    driver.quit()

print("\nCompleted")