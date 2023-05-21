# Pre-requirement modules = pywin32, re, pandas and datetime
# Python script to get and parse Fortigate's IPS email alerts
# This script is looking for a unique source ip address and get all related to him attacks and their dates&times.
# This script also collects the source country of the attacker ip address and the sassionid in forigate logs.
# After data collection is done it then exports only the unique values to Excel file for future analysis.
# Wrote by J.S
# Enjoy

import win32com.client
import re
import pandas as pd
from datetime import datetime


def extract_email_data():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the desired folder in inbox:
    inbox = outlook.GetDefaultFolder(6).Folders("Blacklist")

    # Get all the items in the folder:
    mails = inbox.Items

    email_data = {}

    # Iterate over each item in the folder
    for mail in mails:
        if mail.Class == 43:  # Check  if the item is a mail item
            body = mail.Body

            # Extract the desired fields from the email body
            date = re.search(r"date=(\d{4}-\d{2}-\d{2})", body)
            time = re.search(r"time=(\d{2}:\d{2}:\d{2})", body)
            srcip = re.search(r"srcip=([\d.]+)", body)
            srccountry = re.search(r'srccountry="([^"]+)"', body)
            sessionid = re.search(r"sessionid=([\d]+)", body)
            attack = re.search(r'attack="([^"]+)"', body)
            dstip = re.search(r"dstip=([\d.]+)", body)

            if srcip:
                srcip_value = srcip.group(1)
                # Create a new entry in the email_data dictionary if the srcip is not already present
                if srcip_value not in email_data:
                    email_data[srcip_value] = {
                        "date": set(),
                        "time": set(),
                        "srccountry": set(),
                        "sessionid": set(),
                        "attack": set(),
                        "dstip": set()
                    }

                # Append the extracted values to the respective lists in the email_data dictionary
                if date:
                    email_data[srcip_value]["date"].add(date.group(1))
                if time:
                    email_data[srcip_value]["time"].add(time.group(1))
                if srccountry:
                    email_data[srcip_value]["srccountry"].add(srccountry.group(1))
                if sessionid:
                    email_data[srcip_value]["sessionid"].add(sessionid.group(1))
                if attack:
                    email_data[srcip_value]["attack"].add(attack.group(1))
                if dstip:
                    email_data[srcip_value]["dstip"].add(dstip.group(1))

    return email_data


# Extract the email data
emails = extract_email_data()

# Create a DataFrame from the extracted data
data = []
for srcip, values in emails.items():
    data.append([
        srcip,
        ', '.join(values['srccountry']),
        ', '.join(values['date']),
        ', '.join(values['time']),
        ', '.join(values['sessionid']),
        ', '.join(values['attack']),
        ', '.join(values['dstip'])
    ])
df = pd.DataFrame(data, columns=["Source IP", "Source Country", "Date", "Time", "Session ID", "Attack Method",
                                 "Destination IP"])

# Generate the current date
current_date = datetime.now().strftime("%d-%m-%y")

# Save the DataFrame to an Excel file
excel_filename = str(current_date) + "_FortiGate_Blacklist_Unique.xlsx"
df.to_excel(excel_filename, index=False)
print(f"Email data extracted and saved to {excel_filename} successfully.")
