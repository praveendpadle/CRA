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
            f"window.scrollTo(0,700)"
        )

        #scroll_page(driver)

        # --------------------
# ADDRESS
# --------------------
        address = ""

        try:

            address_container = driver.find_element(
                By.XPATH,
                '//div[@itemprop="address"]'
            )

            address_parts = address_container.find_elements(
                By.XPATH,
                './/span[@itemprop]'
            )

            values = []

            for part in address_parts:

                text = part.text.strip()

                if text:
                    values.append(text)

            address = " ".join(values)

        except:
            pass


        # --------------------
        # PHONE
        # --------------------
        try:

            phone = driver.find_element(
                By.XPATH,
                '//span[@itemprop="telephone"]'
            ).text.strip()

        except:
            pass


        # --------------------
        # HOURS
        # --------------------
        try:
            hours_elem = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span[itemprop='openingHours']")
            )
            )

            hours = hours_elem.text.strip()

            # hour_rows = driver.find_elements(
            #     By.XPATH,
            #     '//tr[contains(@class,"hours")]'
            # )

            # result = []

            # for row in hour_rows:

            #     cols = row.find_elements(
            #         By.XPATH,
            #         './/*[self::td or self::th]'
            #     )

            #     vals = [
            #         c.text.strip()
            #         for c in cols
            #         if c.text.strip()
            #     ]

            #     if len(vals) >= 2:

            #         result.append(
            #             f"{vals[0]}: {vals[1]}"
            #         )

            # if not result:

            #     section = driver.find_element(
            #         By.XPATH,
            #         '//*[contains(text(),"Hours")]'
            #     )

            #     result.append(
            #         section.find_element(
            #             By.XPATH,
            #             './following::*[1]'
            #         ).text
            #     )

            # hours = " | ".join(result)

        except:
            pass


    except Exception as ex:

        error = str(ex)

    return phone, address, hours, error


# -----------------------------------
# Driver
# -----------------------------------
options = uc.ChromeOptions()

options.add_argument("--start-maximized")

driver = uc.Chrome(
    options=options,
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


finally:
    pass
    #driver.quit()


print("\nCompleted")