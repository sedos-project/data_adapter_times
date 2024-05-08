import requests
import pandas as pd


def fetch_data(url):

    response = requests.get(url)
    data = response.json()

    # Convert the JSON data into a DataFrame
    df = pd.DataFrame(data)

    return df


# Fetch Data:
url = "https://openenergy-platform.org/api/v0/schema/model_draft/tables/ind_steel_blafu_0/rows"
fetched_data = fetch_data(url)
print(fetched_data)
