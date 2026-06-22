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
# Append row
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


# -----------------------------------
# Scroll page
# -----------------------------------
def scroll_page(driver):

    positions = [
        300,
        700,
        1200,
        1700
    ]

    for pos in positions:

        driver.execute_script(
            f"window.scrollTo(0,{pos})"
        )

        time.sleep(1)


# -----------------------------------
# Scrape Enterprise
# -----------------------------------
def scrape_enterprise(driver, url):

    address = ""
    phone = ""
    hours = ""
    error = ""

    try:

        driver.get(url)
        print(url)
        wait = WebDriverWait(
            driver,
            30
        )

        wait.until(
            EC.presence_of_element_located(
                (
                    By.TAG_NAME,
                    "body"
                )
            )
        )

        time.sleep(4)

        scroll_page(driver)

        # ----------------
        # HOURS TABLE
        # ----------------
        try:

            hours_table = wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "table.availability-datatable"
                    )
                )
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                hours_table
            )

            time.sleep(2)

        except:
            pass


        # ----------------
        # ADDRESS
        # ----------------
        try:

            addr = driver.find_element(
                By.CSS_SELECTOR,
                "div.location-map-address p.smaller-paragraph"
            )

            address = addr.text.strip()

            print("Address:", address)

        except:
            pass

        # ----------------
        # PHONE
        # ----------------
        try:

            phone = driver.find_element(
                By.XPATH,
                "//a[contains(@href,'tel:')]"
            ).text.strip()

            print("Phone:", phone)

        except:
            pass

        # ----------------
        # HOURS
        # ----------------
        try:

            hrs = []

            rows = driver.find_elements(
                By.CSS_SELECTOR,
                "table.availability-datatable tbody tr"
            )

            for row in rows:

                day = row.find_element(
                    By.TAG_NAME,
                    "th"
                ).text.strip()

                val = row.find_element(
                    By.CSS_SELECTOR,
                    ".location-hour"
                ).text.strip()

                hrs.append(
                    f"{day}: {val}"
                )

            hours = " | ".join(
                hrs
            )

            print("Hours:", hours)

        except:
            pass

    except Exception as ex:

        error = str(ex)

    return address, phone, hours, error


# -----------------------------------
# DRIVER
# -----------------------------------
options = uc.ChromeOptions()

options.add_argument(
    "--start-maximized"
)

driver = uc.Chrome(
    options=options,
    version_main=149,
    use_subprocess=True
)


# -----------------------------------
# INPUT
# -----------------------------------
df = pd.read_excel(
    INPUT_FILE
)


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

        print(
            f"\nProcessing → {featureid}"
        )

        address, phone, hours, error = (
            scrape_enterprise(
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

        append_excel(
            output
        )

        print("Saved")

        time.sleep(2)

finally:

    #driver.quit()
    pass


print("\nCompleted")