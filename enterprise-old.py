from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def scrape_enterprise_atl(url):
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless=new")  # run headless if needed
    driver = webdriver.Chrome(options=options)
    driver.get(url)

    wait = WebDriverWait(driver, 20)

    # --- scroll to the hours section (force lazy loading) ---
    try:
        hours_table = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.availability-datatable"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", hours_table)
        time.sleep(2)  # give lazy loader time to populate
    except Exception as e:
        print("Hours table not found:", e)

    data = {}

    # --- Address ---
    try:
        addr_elem = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.location-map-address p.smaller-paragraph"))
        )
        data["address"] = addr_elem.text.strip()
    except:
        data["address"] = None

    # --- Phone ---
    try:
        phone_elem = driver.find_element(By.XPATH, "//a[starts-with(@href, 'tel:')]")
        data["phone"] = phone_elem.text.strip()
    except:
        data["phone"] = None

    # --- Opening Hours ---
    hours = {}
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "table.availability-datatable tbody tr")
        for row in rows:
            day = row.find_element(By.TAG_NAME, "th").text.strip()
            timing = row.find_element(By.CSS_SELECTOR, ".location-hour").text.strip()
            hours[day] = timing
        data["hours"] = hours
    except Exception as e:
        print("Error extracting hours:", e)
        data["hours"] = None

    driver.quit()
    return data


if __name__ == "__main__":
    url =  "https://www.enterprise.com/en/car-rental-locations/us/ga/atlanta-hartsfield-jackson-intl-airport-030e.html?mcid=yext:245709"
    result = scrape_enterprise_atl(url)
    print(result)
