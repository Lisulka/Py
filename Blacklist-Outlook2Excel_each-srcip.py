# Pre-requirement modules = pywin32, re, pandas and datetime
# Python script to get and parse Fortigate's IPS email alerts
# This script is looking for a unique source ip address and get all related to him attacks and their dates&times.
# This script also collects the source country of the attacker ip address and the sassionid in forigate logs.
# After data collection is done it then exports to separate Excel file for every srcip for future analysis.
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
                        "date": [],
                        "time": [],
                        "srccountry": [],
                        "sessionid": [],
                        "attack": [],
                        "dstip": []
                    }
                # Append the extracted values to the respective lists in the email_data dictionary
                if date:
                    email_data[srcip_value]["date"].append(date.group(1))
                if time:
                    email_data[srcip_value]["time"].append(time.group(1))
                if srccountry:
                    email_data[srcip_value]["srccountry"].append(srccountry.group(1))
                if sessionid:
                    email_data[srcip_value]["sessionid"].append(sessionid.group(1))
                if attack:
                    email_data[srcip_value]["attack"].append(attack.group(1))
                if dstip:
                    email_data[srcip_value]["dstip"].append(dstip.group(1))

    return email_data


# Extract the email data
emails = extract_email_data()

# Generate the current date
current_date = datetime.now().strftime("%d-%m-%y")

# Create separate files for each srcip
for srcip, values in emails.items():
    data = {
        "Source Country": values["srccountry"],
        "Session ID": values["sessionid"],
        "Date": values["date"],
        "Time": values["time"],
        "Attack Method": values["attack"],
        "Destination IP": values["dstip"]
    }
    df = pd.DataFrame(data)

    # Save the DataFrame to a separate Excel file for each srcip
    excel_filename = str(srcip) + str("_") + str(current_date) + "_FortiGate_Blacklist.xlsx"
    df.to_excel(excel_filename, index=False)
    print(f"Email data for {srcip} extracted and saved to {excel_filename} successfully.")
