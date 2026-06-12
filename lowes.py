import os
import time
import pandas as pd
from curl_cffi import requests
from openpyxl import load_workbook


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "lowes_output.xlsx"

BASE_URL = (
    "https://www.lowes.com/store/api/search"
)

session = requests.Session(
    impersonate="chrome"
)


# -----------------------
# Format Hours
# -----------------------
def format_hours(hours):

    result = []

    for h in hours:

        day = h.get(
            "day",
            {}
        )

        result.append(
            f"{day.get('day')} "
            f"{day.get('open')}-"
            f"{day.get('close')}"
        )

    return " | ".join(result)


# -----------------------
# Append to Excel
# -----------------------
def append_excel(rows):

    if not rows:
        return

    df = pd.DataFrame(rows)

    if not os.path.exists(
        OUTPUT_FILE
    ):

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

            workbook = writer.book

            sheet = workbook.active

            startrow = (
                sheet.max_row
            )

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=startrow
            )


# -----------------------
# Extract JSON
# -----------------------
def extract(data):

    rows = []

    for item in data.get("stores", []):

        store = item.get(
            "store",
            {}
        ).copy()

        # Add distance from parent
        store["distance"] = (
            item.get(
                "distance"
            )
        )

        # Convert storeHours list
        if "storeHours" in store:

            hours = []

            for h in store[
                "storeHours"
            ]:

                day = h.get(
                    "day",
                    {}
                )

                hours.append(
                    f"{day.get('day')} "
                    f"{day.get('open')}-"
                    f"{day.get('close')}"
                )

            store[
                "storeHours"
            ] = " | ".join(
                hours
            )

        # Generate store URL
        state = (
            str(
                store.get(
                    "state",
                    ""
                )
            )
            .strip()
        )

        city = (
            str(
                store.get(
                    "city",
                    ""
                )
            )
            .strip()
            .replace(
                " ",
                "-"
            )
        )

        store_id = (
            str(
                store.get(
                    "id",
                    ""
                )
            )
            .strip()
        )

        store[
            "store_url"
        ] = (
            "https://www.lowes.com/store/"
            f"{state}-{city}/"
            f"{store_id}"
        )

        rows.append(
            store
        )

    return rows

# -----------------------
# Fetch API
# -----------------------
def fetch(zipcode):

    params = {

        "responseGroup":
            "large",

        "searchTerm":
            str(
                zipcode
            )
    }

    response = session.get(
        BASE_URL,
        params=params,
        timeout=60
    )

    response.raise_for_status()

    return response.json()


# -----------------------
# Main
# -----------------------
# -----------------------
# Main
# -----------------------
def process():

    input_df = pd.read_excel(
        INPUT_FILE
    )

    for _, row in input_df.iterrows():

        zipcode = str(
            row["Pincode"]
        )

        if len(zipcode)== 3:
            zipcode = zipcode.zfill(5)

        if len(zipcode)== 4:
            zipcode = zipcode.zfill(5)
       
        try:

            print(
                f"Processing {zipcode}"
            )

            data = fetch(
                zipcode
            )

            rows = extract(
                data
            )

            append_excel(
                rows
            )

            print(
                f"Saved {len(rows)} rows"
            )

            time.sleep(1)

        except Exception as e:

            print(
                f"Error {zipcode}: {e}")


process()

print("Completed")
 