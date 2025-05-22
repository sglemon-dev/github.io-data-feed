import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import json
import os

# Read Excel file
df = pd.read_excel('PPS_Brevo_Data_Feed_Customer_Purchases.xlsx', sheet_name='pps_customer_orders')

# Clean and convert date columns to ISO format
df['Order Date'] = pd.to_datetime(df['Order Date']).dt.strftime('%Y-%m-%dT%H:%M:%SZ')
df['Date Created'] = pd.to_datetime(df['Date Created']).dt.strftime('%Y-%m-%dT%H:%M:%SZ')

# Convert to JSON
json_data = df.to_dict(orient='records')
with open('data-feed.json', 'w', encoding='utf-8') as f:
    json.dump(json_data, f, indent=2, ensure_ascii=False)

# Create RSS feed
rss = ET.Element('rss', version='2.0')
channel = ET.SubElement(rss, 'channel')
ET.SubElement(channel, 'title').text = 'PPS Customer Purchases Feed'
ET.SubElement(channel, 'link').text = 'https://example.com/data-feed.xml'
ET.SubElement(channel, 'description').text = 'Customer purchase data from PPS'
ET.SubElement(channel, 'lastBuildDate').text = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')

for _, row in df.iterrows():
    item = ET.SubElement(channel, 'item')
    ET.SubElement(item, 'title').text = f"Purchase: {row['Description']} by {row['First Name']} {row['Last Name']}"
    ET.SubElement(item, 'link').text = f"https://example.com/order/{row['ID Number']}"
    ET.SubElement(item, 'guid', isPermaLink='false').text = f"order-{row['ID Number']}"
    ET.SubElement(item, 'pubDate').text = datetime.strptime(row['Order Date'], '%Y-%m-%dT%H:%M:%SZ').strftime('%a, %d %b %Y %H:%M:%S GMT')
    description = (
        f"Customer ID: {row['Customer ID']}<br>"
        f"Email: {row['Email']}<br>"
        f"Item Code: {row['Item Code']}<br>"
        f"Description: {row['Description']}<br>"
        f"Order Date: {row['Order Date']}<br>"
        f"First Name: {row['First Name']}<br>"
        f"Last Name: {row['Last Name']}<br>"
        f"Date Created: {row['Date Created']}"
    )
    ET.SubElement(item, 'description').text = description

# Write RSS to XML file
tree = ET.ElementTree(rss)
with open('data-feed.xml', 'wb') as f:
    tree.write(f, encoding='utf-8', xml_declaration=True)

print("Generated data-feed.json and data-feed.xml")