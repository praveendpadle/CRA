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
# Scroll page for lazy loading
# -----------------------------------
def scroll_page(driver):

    positions = [500, 1000, 1500, 2000]

    for pos in positions:

        driver.execute_script(
            f"window.scrollTo(0,{pos})"
        )

        time.sleep(2)

    driver.execute_script(
        "window.scrollTo(0,0)"
    )

    time.sleep(1)


# -----------------------------------
# Append to Excel
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

            start_row = writer.book.active.max_row

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=start_row
            )


# -----------------------------------
# Hertz Scraper
# -----------------------------------
def scrape_hertz(driver, url):

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

        time.sleep(5)

        scroll_page(driver)

        # --------------------
        # ADDRESS
        # --------------------
        address = ""

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
                    line.text.strip()
                    for line in lines
                    if line.text.strip()
                ]
            )

            print("Address:", address)

        except Exception as ex:

            print("Address Error:", ex)


        # --------------------
        # PHONE
        # --------------------
        try:

            phone = driver.find_element(
                By.XPATH,
                '//a[contains(@href,"tel:")]'
            ).text.strip()

            print("Phone:", phone)

        except Exception as ex:

            print("Phone Error:", ex)


        # --------------------
        # OPERATING HOURS
        # --------------------
        try:

            operating_hours = []

            hours_section = driver.find_element(
                By.XPATH,
                '//h3[contains(text(),"Hours")]/parent::*'
            )

            rows = hours_section.find_elements(
                By.XPATH,
                './/div[contains(@class,"70qvj9")]'
            )

            for row in rows:

                values = row.find_elements(
                    By.TAG_NAME,
                    "p"
                )

                if len(values) >= 2:

                    day = values[0].text.strip()
                    timing = values[1].text.strip()

                    operating_hours.append(
                        f"{day}: {timing}"
                    )

            hours = " | ".join(
                operating_hours
            )

            print("Hours:", hours)

        except Exception as ex:

            print("Hours Error:", ex)

            # fallback
            try:

                rows = driver.find_elements(
                    By.XPATH,
                    '//h3[contains(text(),"Hours")]//following::div[1]//p'
                )

                temp = []

                for i in range(0, len(rows), 2):

                    try:

                        day = rows[i].text.strip()
                        timing = rows[i + 1].text.strip()

                        temp.append(
                            f"{day}: {timing}"
                        )

                    except:
                        pass

                hours = " | ".join(temp)

            except:
                pass

    except Exception as ex:

        error = str(ex)

    return phone, address, hours, error


# -----------------------------------
# Chrome Driver
# -----------------------------------
options = uc.ChromeOptions()

options.add_argument("--start-maximized")
#options.add_argument('--headless')
driver = uc.Chrome(
    options=options,
    version_main=149,      # Change if your Chrome version differs
    use_subprocess=True
)


# -----------------------------------
# Read Input Excel
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

        print(f"\nProcessing -> {featureid}")

        phone, address, hours, error = scrape_hertz(
            driver,
            website
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

    driver.quit()


print("\nCompleted")