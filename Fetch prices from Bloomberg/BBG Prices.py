import pandas as pd
import requests

# Load Excel file containing URLs
df = pd.read_excel('api_urls.xlsx')  # Replace with your Excel file

data_list = []
timeout = 1000  # Timeout in seconds

for url in df.iloc[:, 0]:  # Assuming URLs are in the first column
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()  # Raise exception for HTTP errors

        if response.status_code == 200:
            data = response.json()

            # Extract field names and values
            field_names = data['fields']['field']
            values = [entry['value'] for entry in data['instrumentDatas']['instrumentData'][0]['data']]
            data_dict = dict(zip(field_names, values))
            data_list.append(data_dict)

            # Log the fetched data
            for field, value in data_dict.items():
                print(f"{field}: {value}")
    except Exception as e:
        print(f"Failed to fetch data from {url}: {e}")

# Combine original URLs and fetched data into a DataFrame
output_df = pd.concat([df, pd.DataFrame(data_list)], axis=1)

# Save to Excel
output_df.to_excel('output_data.xlsx', index=False)
print("Data saved to output_data.xlsx file.")
