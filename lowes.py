import os
import time
import pandas as pd
from curl_cffi import requests


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "lowes_output.xlsx"

BASE_URL = (
    "https://www.lowes.com/store/api/search"
)

session = requests.Session(
    impersonate="chrome"
)

store_cache = {}


# ---------------------------------
# Load Existing IDs
# ---------------------------------
def load_existing_ids():

    if not os.path.exists(
        OUTPUT_FILE
    ):
        return

    try:

        df = pd.read_excel(
            OUTPUT_FILE,
            usecols=["id"]
        )

        for x in (
            df["id"]
            .dropna()
            .astype(str)
        ):

            store_cache[x] = True

        print(
            f"Loaded "
            f"{len(store_cache)} "
            f"existing IDs"
        )

    except:
        pass


# ---------------------------------
# Hours Formatter
# ---------------------------------
def format_hours(hours):

    result = []

    for h in hours:

        d = h.get(
            "day",
            {}
        )

        result.append(
            f"{d.get('day')} "
            f"{d.get('open')}-"
            f"{d.get('close')}"
        )

    return (
        " | ".join(
            result
        )
    )


# ---------------------------------
# Extract Fields
# ---------------------------------
def extract(
    zipcode,
    data
):

    rows = []

    for item in (
        data.get(
            "stores",
            []
        )
    ):

        s = item.get(
            "store",
            {}
        )

        rows.append({

            "inputzip":
                zipcode,

            "_id":
                s.get(
                    "_id"
                ),

            "zip":
                s.get(
                    "zip"
                ),

            "id":
                s.get(
                    "id"
                ),

            "bis_name":
                s.get(
                    "bis_name"
                ),

            "store_name":
                s.get(
                    "store_name"
                ),

            "address":
                s.get(
                    "address"
                ),

            "storeHours":
                format_hours(
                    s.get(
                        "storeHours",
                        []
                    )
                ),

            "city":
                s.get(
                    "city"
                ),

            "state":
                s.get(
                    "state"
                ),

            "phone":
                s.get(
                    "phone"
                ),

            "fax":
                s.get(
                    "fax"
                ),

            "proServicesDesk":
                s.get(
                    "proServicesDesk"
                ),

            "proFax":
                s.get(
                    "proFax"
                ),

            "lat":
                s.get(
                    "lat"
                ),

            "long":
                s.get(
                    "long"
                ),

            "areaNumber":
                s.get(
                    "areaNumber"
                ),

            "storeType":
                s.get(
                    "storeType"
                ),

            "storeStatusCd":
                s.get(
                    "storeStatusCd"
                ),

            "openDate":
                s.get(
                    "openDate"
                )
        })

    return rows


# ---------------------------------
# Save Incrementally
# ---------------------------------
def append_excel(rows):

    unique = []

    for row in rows:

        sid = str(
            row["id"]
        )

        if sid in store_cache:

            print(
                f"Skip {sid}"
            )

            continue

        store_cache[
            sid
        ] = True

        unique.append(
            row
        )

    if not unique:
        return

    df = pd.DataFrame(
        unique
    )

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

            ws = (
                writer.book.active
            )

            start = (
                ws.max_row
            )

            df.to_excel(
                writer,
                index=False,
                header=False,
                startrow=start
            )

    print(
        f"Saved "
        f"{len(unique)}"
    )


# ---------------------------------
# API Call
# ---------------------------------
def fetch(
    zipcode
):

    params = {

        "responseGroup":
            "large",

        "searchTerm":
            zipcode
    }

    r = session.get(
        BASE_URL,
        params=params,
        timeout=60
    )

    r.raise_for_status()

    return r.json()


# ---------------------------------
# Process
# ---------------------------------
def process():

    input_df = pd.read_excel(
        INPUT_FILE
    )

    for _, row in (
        input_df.iterrows()
    ):

        zipcode = str(
            row[
                "Pincode"
            ]
        )

        try:

            print(
                f"\nProcessing "
                f"{zipcode}"
            )

            data = fetch(
                zipcode
            )

            rows = extract(
                zipcode,
                data
            )

            append_excel(
                rows
            )

            time.sleep(1)

        except Exception as e:

            print(
                f"Error "
                f"{zipcode}"
            )

            print(e)


# ---------------------------------
# Run
# ---------------------------------
load_existing_ids()

process()

print(
    "\nCompleted"
)