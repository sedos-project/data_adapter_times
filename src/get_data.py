import requests
import pandas as pd

# Fetch the data
url = "https://openenergy-platform.org/api/v0/schema/model_draft/tables/ind_steel_blafu_0/rows"
response = requests.get(url)
data = response.json()

# Convert the JSON data into a DataFrame
df = pd.DataFrame(data)

# Display the first few rows of the DataFrame
print(df.head())
