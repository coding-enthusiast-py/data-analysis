# Bloomberg Price Fetcher

This Python script fetches pricing data from Bloomberg API endpoints stored in an Excel file. It is intended for general usage where JSON responses include Bloomberg-style structured data (e.g., fields and instrument data).

# Requirements

- Python 3.7+
- pandas
- requests
- openpyxl (for Excel file reading/writing)

Install dependencies with:

```bash
pip install pandas requests openpyxl

# Input
An Excel file (api_urls.xlsx) containing a list of API URLs (one per row). Example:

# API URLs
https://api.example.com/bbg/abc123
https://api.example.com/bbg/xyz789

Ensure the URLs return JSON responses in the following structure:

json
Copy
Edit
{
  "fields": {
    "field": ["PX_LAST", "BID", "ASK"]
  },
  "instrumentDatas": {
    "instrumentData": [{
      "data": [
        {"value": 100.5},
        {"value": 100.2},
        {"value": 100.8}
      ]
    }]
  }
}


# Output

After fetching the data, the script:
Parses the fields and associated values
Combines them with the input URLs
Exports a consolidated Excel file output_data.xlsx
