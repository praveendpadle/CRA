import pandas as pd
import time
from curl_cffi import requests

INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "tmobile_output.xlsx"

BASE_URL = "https://www.t-mobile.com/stores/api/get-nearby-business"

session = requests.Session(
    impersonate="chrome"
)

results = []

df = pd.read_excel(INPUT_FILE)

for _, row in df.iterrows():

    pincode = str(row["Pincode"]).strip()
    state = str(row["State"]).strip()
    country = str(row["Country"]).strip()

    lat = row["Lat"]
    lng = row["Lng"]

    page = 1

    loc = f"{state} {pincode}, {country}"

    print(f"\nProcessing {loc}")

    while True:

        params = {
            "INTNAV": "tNav:StoreLocator",
            "page": page,
            "loc": loc,
            "lat": lat,
            "lon": lng,
            "bsort": "recommended"
        }

        try:

            response = session.get(
                BASE_URL,
                params=params,
                headers={
                    "Referer":
                        "https://www.t-mobile.com/stores",
                    "Accept":
                        "application/json"
                },
                timeout=60
            )

            print(
                "Status:",
                response.status_code
            )

            response.raise_for_status()

            data = response.json()

            business = data.get(
                "business_list",
                {}
            )

            stores = business.get(
                "object_list",
                []
            )

            if not stores:
                break

            for item in stores:

                results.append({

                    "InputPincode":
                        pincode,

                    "display_name":
                        item.get(
                            "display_name"
                        ),

                    "external_store_code":
                        item.get(
                            "external_store_code"
                        ),

                    "lat":
                        item.get(
                            "lat"
                        ),

                    "lon":
                        item.get(
                            "lon"
                        ),

                    "address_postcode":
                        item.get(
                            "address_postcode"
                        ),

                    "schemaHrs":
                        ", ".join(
                            item.get(
                                "all_opening_hours",
                                {}
                            ).get(
                                "schemaHrs",
                                []
                            )
                        ),

                    "product_link":
                        item.get(
                            "product_link"
                        ),

                    "business_phone_text":
                        item.get(
                            "contact_context",
                            {}
                        ).get(
                            "business_phone_text"
                        ),

                    "address_text":
                        item.get(
                            "address_text"
                        ),

                    "get_directions_link":
                        item.get(
                            "get_directions_link"
                        )
                })
            next_page = business.get(
                "next_page_number"
            )

            print(
                "Next Page:",
                next_page
            )

            if next_page is None:
                break

            page = next_page

            time.sleep(1)

        except Exception as e:

            print(
                f"Failed: {e}"
            )

            break


pd.DataFrame(results).to_excel(
    OUTPUT_FILE,
    index=False
)

print("\nDone")