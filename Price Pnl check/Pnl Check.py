import os
import re
import json
import tabula
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import requests
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
import win32com.client

# Setup date references
today = datetime.today().date()

if today.weekday() == 0:
    previous_weekday = today - timedelta(days=3)
    if previous_weekday.weekday() == 0:
        previous_to_previous_weekday = previous_weekday - timedelta(days=3)
    else:
        previous_to_previous_weekday = previous_weekday - timedelta(days=1)
else:
    previous_weekday = today - timedelta(days=1)
    if previous_weekday.weekday() == 0:
        previous_to_previous_weekday = previous_weekday - timedelta(days=3)
    else:
        previous_to_previous_weekday = previous_weekday - timedelta(days=1)

east_date_format = previous_weekday.strftime("%d %B %Y")
drake_date_format = previous_weekday.strftime("%d-%m-%y")
previous_weekday_formatted = previous_weekday.strftime("%Y-%m-%d")
previous_to_previous_weekday_formatted = previous_to_previous_weekday.strftime("%Y-%m-%d")

# Access Outlook for mail extraction
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
root = outlook.Folders.Item("your_email@yourdomain.com")  # Replace with your email
deleted_folder = root.Folders.Item("Deleted Items")

# Look for files in mail
for msg in deleted_folder.Items:
    if "PnL Check for " + previous_weekday_formatted in str(msg.Subject):
        for attachment in msg.Attachments:
            path = f'C://local//PnL_Check_{previous_weekday_formatted}.xlsx'
            attachment.SaveAsFile(path)
            print(f"File saved: {path}")

# Replace the logic as needed based on your folder structure
inbox_folder = root.Folders.Item("Inbox")
custom_folder1 = inbox_folder.Folders.Item("Team_A")
custom_folder2 = custom_folder1.Folders.Item("Pricing_Review")

latest_message = None
mail_items = custom_folder2.Items
mail_items.Sort("[ReceivedTime]", True)

for mail in mail_items:
    if f"Performance Report - {east_date_format}" in str(mail.Subject):
        latest_message = mail
        break

if latest_message:
    for att in latest_message.Attachments:
        file_path = f'C://local//TeamA_Report_{previous_weekday_formatted}.csv'
        att.SaveAsFile(file_path)
        print(f"Saved: {file_path}")
        df_east = pd.read_csv(file_path)
        expected_DTD = df_east['TotalPl'].sum()
        print(expected_DTD)

# Load CAM Excel data
df = pd.read_excel(f'C:\\local\\PnL_Check_{previous_weekday_formatted}.xlsx')
df['Investment'] = pd.to_numeric(df['Investment'], errors='coerce').fillna(df['Investment']).astype(object)
df['comment'] = ''

# Filter by business logic
condition1 = (df['PnL_Greater_Less_then_10000'] != 0) & df['PnL_Greater_Less_then_10000'].notnull()
condition2 = abs(df['PnL_Impact_bps']) >= 25

filtered_df = df[condition1 & condition2 & (df['Subtype'] != 'Fund of Funds')]
filtered_df3 = df[condition1 & condition2]

# API calls for pricing and comments
all_tables = []
for index, row in filtered_df.iterrows():
    try:
        spn = int(row[0])
        url = f'https://api-base-url.com/service/securityService/getSecurities?spns=[{spn}]&format=json'
        kerb_auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL)
        response = requests.get(url, auth=kerb_auth)

        if response.status_code == 200:
            data_json = response.json()
            fields = data_json["fields"]
            data = data_json["data"]
            column_names = [field["name"] for field in fields]
            canonical = pd.DataFrame(data, columns=column_names)
            canonical = canonical[["spn", "canonicalSpn", "subtypeId"]]

            if not canonical.empty and 'canonicalSpn' in canonical.columns and (canonical['subtypeId'] != 32).any():
                spn = str(canonical['canonicalSpn'].iloc[0])
                df.loc[index, 'comment'] = 'Canonical return: '

        pricing_url = f'https://api-base-url.com/pricing/fetchQuoteDetails.action?spn={spn}&date={previous_weekday_formatted}&pricingTypeCode=EQ'
        response = requests.get(pricing_url, auth=kerb_auth)

        if response.status_code == 200:
            tables = pd.read_html(response.content)
            if tables:
                meta = {
                    'SPN': tables[0][0].str.split(": ").str[-1].iat[0],
                    'Date': tables[0][1].str.split(": ").str[-1].iat[0],
                    'Desname': tables[0][2].str.split(": ").str[-1].iat[0]
                }

                for table in tables[1:]:
                    table['SPN'] = meta['SPN']
                    table['Date'] = meta['Date']
                    table['Desname'] = meta['Desname']
                    table = table.drop(table[(table['Price'] == 0) & (table['Source'] == 'Third Party Source')].index)
                    if len(table) > 1 and table.duplicated(subset=['Price']).any():
                        df.loc[index, 'comment'] += 'Price discrepancy detected from agreed hierarchy'
                    elif (table['Source'] == 'Preferred Source1').any():
                        df.loc[index, 'comment'] += 'Price picked from preferred_source1'

                all_tables.extend(tables[1:])
            else:
                print("No tables found")

    except Exception as e:
        print(f"Error processing SPN {spn}: {e}")

# Output results
merged_table = pd.concat(all_tables, ignore_index=True)
merged_table.to_excel("C:\\local\\Pricing_Source_Data.xlsx", index=False)

# Add final formatting
df.to_excel(f'C:\\local\\PnL_Check_{previous_weekday_formatted}.xlsx', index=False)
