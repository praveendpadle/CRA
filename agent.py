import os
from google.adk.agents.llm_agent import Agent

import re
import requests
from bs4 import BeautifulSoup
import time

from selenium import webdriver

from selenium.webdriver.chrome.service import Service

from selenium.webdriver.chrome.options import Options

from webdriver_manager.chrome import ChromeDriverManager

from google.adk.agents import Agent
#from tools.scraper_tool import scrape_business_details

def scrape_business_details_Old(url: str):

    headers = {

        "User-Agent": "Mozilla/5.0"

    }

    try:

        response = requests.get(

            url,

            headers=headers,

            timeout=20

        )

        response.raise_for_status()

    except Exception as e:

        return {

            "url": url,

            "error": str(e)

        }

    soup = BeautifulSoup(response.text, "lxml")

    text = soup.get_text(" ", strip=True)

    # ==========================================

    # PHONE NUMBER EXTRACTION

    # ==========================================

    phone_pattern = re.compile(

        r'(\+?\d[\d\s\-\(\)]{8,}\d)'

    )

    phones = list(set(phone_pattern.findall(text)))

    # Clean numbers

    cleaned_phones = []

    for p in phones:

        p = re.sub(r'\s+', ' ', p).strip()

        if len(p) >= 10:

            cleaned_phones.append(p)

    phone_number = cleaned_phones[0] if cleaned_phones else ""

    # ==========================================

    # ADDRESS EXTRACTION

    # ==========================================

    address = ""

    address_keywords = [

        "address",

        "location",

        "head office",

        "corporate office",

        "registered office",

        "contact us"

    ]

    for tag in soup.find_all(["p", "div", "span", "li"]):

        content = tag.get_text(" ", strip=True)

        if any(

            keyword.lower() in content.lower()

            for keyword in address_keywords

        ):

            address = content

            break

    # ==========================================

    # OPERATING HOURS EXTRACTION

    # ==========================================

    operating_hours = {}

    days_map = {

        "mon": "Monday",

        "tue": "Tuesday",

        "wed": "Wednesday",

        "thu": "Thursday",

        "fri": "Friday",

        "sat": "Saturday",

        "sun": "Sunday"

    }

    # ------------------------------------------

    # METHOD 1: schema.org itemprop

    # ------------------------------------------

    possible_hours = soup.find_all(

        attrs={"itemprop": "openingHours"}

    )

    for h in possible_hours:

        val = h.get_text(" ", strip=True)

        if not val:

            continue

        # Example:

        # Mon - Fri 9:00 AM - 8:00 PM

        # Sun 10:00 AM - 5:00 PM

        pattern = re.compile(

            r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

            r'(?:\s*-\s*'

            r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun))?'

            r'[:\s]*'

            r'(.+)',

            re.IGNORECASE

        )

        match = pattern.search(val)

        if match:

            start_day = match.group(1)

            end_day = match.group(2)

            timing = match.group(3).strip()

            day_keys = list(days_map.keys())

            start_index = day_keys.index(

                start_day[:3].lower()

            )

            if end_day:

                end_index = day_keys.index(

                    end_day[:3].lower()

                )

            else:

                end_index = start_index

            for i in range(start_index, end_index + 1):

                full_day = days_map[

                    day_keys[i]

                ]

                operating_hours[

                    full_day

                ] = timing

    # ------------------------------------------

    # METHOD 2: REGEX FALLBACK

    # ------------------------------------------

    if not operating_hours:

        patterns = [

            # Monday: 9 AM - 9 PM

            r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)'

            r'\s*[:\-]?\s*'

            r'([0-9:\sAMPMapm\-\–toTOClosedclosed]+)',

            # Mon: 9 AM - 9 PM

            r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

            r'\s*[:\-]?\s*'

            r'([0-9:\sAMPMapm\-\–toTOClosedclosed]+)'

        ]

        for pattern in patterns:

            matches = re.findall(

                pattern,

                text,

                re.IGNORECASE

            )

            for day, timing in matches:

                short_day = day[:3].lower()

                if short_day in days_map:

                    full_day = days_map[short_day]

                    operating_hours[

                        full_day

                    ] = timing.strip()

    # ------------------------------------------

    # METHOD 3: HANDLE DAY RANGES

    # Example:

    # Mon to Fri : 9AM - 9PM

    # ------------------------------------------

    if not operating_hours:

        range_pattern = re.compile(

            r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

            r'\s*(?:to|-)\s*'

            r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

            r'\s*[:\-]?\s*'

            r'([0-9:\sAMPMapm\-\–]+)',

            re.IGNORECASE

        )

        matches = range_pattern.findall(text)

        day_keys = list(days_map.keys())

        for start_day, end_day, timing in matches:

            start_index = day_keys.index(

                start_day[:3].lower()

            )

            end_index = day_keys.index(

                end_day[:3].lower()

            )

            for i in range(start_index, end_index + 1):

                full_day = days_map[

                    day_keys[i]

                ]

                operating_hours[

                    full_day

                ] = timing.strip()

    # ==========================================

    # RETURN RESULT

    # ==========================================

    return {

        "url": url,

        "address": address,

        "phone_number": phone_number,

        "operating_hours": operating_hours

    }

def scrape_business_details(url: str):

    # ==========================================

    # CHROME OPTIONS

    # ==========================================

    chrome_options = Options()

    #chrome_options.add_argument("--headless")

    chrome_options.add_argument("--disable-gpu")

    chrome_options.add_argument("--no-sandbox")

    chrome_options.add_argument("--disable-dev-shm-usage")

    chrome_options.add_argument(

        "user-agent=Mozilla/5.0"

    )

    # ==========================================

    # START DRIVER

    # ==========================================

    driver = webdriver.Chrome(

        service=Service(

            ChromeDriverManager().install()

        ),

        options=chrome_options

    )

    try:

        # ==========================================

        # OPEN URL

        # ==========================================

        driver.get(url)

        # Wait for JS rendering

        time.sleep(5)

        page_text = driver.find_element(

            "tag name",

            "body"

        ).text

        # ==========================================

        # PHONE EXTRACTION

        # ==========================================

        phone_pattern = re.compile(

            r'(\+?\d[\d\s\-\(\)]{8,}\d)'

        )

        phones = list(

            set(

                phone_pattern.findall(page_text)

            )

        )

        cleaned_phones = []

        for p in phones:

            p = re.sub(

                r'\s+',

                ' ',

                p

            ).strip()

            if len(p) >= 10:

                cleaned_phones.append(p)

        phone_number = (

            cleaned_phones[0]

            if cleaned_phones

            else ""

        )

        # ==========================================

        # ADDRESS EXTRACTION

        # ==========================================

        address = ""

        address_keywords = [

            "address",

            "location",

            "head office",

            "corporate office",

            "registered office",

            "contact us"

        ]

        elements = driver.find_elements(

            "xpath",

            "//*[self::p or self::div or self::span or self::li]"

        )

        for element in elements:

            try:

                content = element.text.strip()

                if any(

                    keyword.lower() in content.lower()

                    for keyword in address_keywords

                ):

                    address = content

                    break

            except:

                pass

        # ==========================================

        # OPERATING HOURS EXTRACTION

        # ==========================================

        operating_hours = {}

        days_map = {

            "mon": "Monday",

            "tue": "Tuesday",

            "wed": "Wednesday",

            "thu": "Thursday",

            "fri": "Friday",

            "sat": "Saturday",

            "sun": "Sunday"

        }

        # ==========================================

        # METHOD 1:

        # itemprop=openingHours

        # ==========================================

        hour_elements = driver.find_elements(

            "xpath",

            '//*[@itemprop="openingHours"]'

        )

        for h in hour_elements:

            try:

                val = h.text.strip()

                if not val:

                    continue

                pattern = re.compile(

                    r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

                    r'(?:\s*-\s*'

                    r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun))?'

                    r'[:\s]*'

                    r'(.+)',

                    re.IGNORECASE

                )

                match = pattern.search(val)

                if match:

                    start_day = match.group(1)

                    end_day = match.group(2)

                    timing = match.group(3).strip()

                    day_keys = list(days_map.keys())

                    start_index = day_keys.index(

                        start_day[:3].lower()

                    )

                    if end_day:

                        end_index = day_keys.index(

                            end_day[:3].lower()

                        )

                    else:

                        end_index = start_index

                    for i in range(

                        start_index,

                        end_index + 1

                    ):

                        full_day = days_map[

                            day_keys[i]

                        ]

                        operating_hours[

                            full_day

                        ] = timing

            except:

                pass

        # ==========================================

        # METHOD 2:

        # FULL PAGE REGEX

        # ==========================================

        if not operating_hours:

            patterns = [

                # Monday: 9 AM - 9 PM

                r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)'

                r'\s*[:\-]?\s*'

                r'([0-9:\sAMPMapm\-\–toTOClosedclosed]+)',

                # Mon: 9 AM - 9 PM

                r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

                r'\s*[:\-]?\s*'

                r'([0-9:\sAMPMapm\-\–toTOClosedclosed]+)'

            ]

            for pattern in patterns:

                matches = re.findall(

                    pattern,

                    page_text,

                    re.IGNORECASE

                )

                for day, timing in matches:

                    short_day = day[:3].lower()

                    if short_day in days_map:

                        full_day = days_map[

                            short_day

                        ]

                        operating_hours[

                            full_day

                        ] = timing.strip()

        # ==========================================

        # METHOD 3:

        # Mon to Fri : 9AM - 9PM

        # ==========================================

        if not operating_hours:

            range_pattern = re.compile(

                r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

                r'\s*(?:to|-)\s*'

                r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)'

                r'\s*[:\-]?\s*'

                r'([0-9:\sAMPMapm\-\–]+)',

                re.IGNORECASE

            )

            matches = range_pattern.findall(

                page_text

            )

            day_keys = list(days_map.keys())

            for start_day, end_day, timing in matches:

                start_index = day_keys.index(

                    start_day[:3].lower()

                )

                end_index = day_keys.index(

                    end_day[:3].lower()

                )

                for i in range(

                    start_index,

                    end_index + 1

                ):

                    full_day = days_map[

                        day_keys[i]

                    ]

                    operating_hours[

                        full_day

                    ] = timing.strip()

        # ==========================================

        # RETURN RESULT

        # ==========================================

        return {

            "url": url,

            "address": address,

            "phone_number": phone_number,

            "operating_hours": operating_hours

        }

    except Exception as e:

        return {

            "url": url,

            "error": str(e)

        }

    finally:

        driver.quit()

root_agent = Agent(
    name="business_scraper_agent",
    model="gemini-3-flash-preview",
    description="Extract business details from URLs",
    instruction="""
You are a web scraping assistant.

Your task:
1. Accept a URL
2. Use scraper tool
3. Extract:
   - Address
   - Phone number
   - Operating hours(eg. Mon to Sun:9AM - 9PM ) consider if hour changes for particular day.
4. Return structured JSON
""",
    tools=[scrape_business_details]
)
 
