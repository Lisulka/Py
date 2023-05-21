# Pre-required modules = pywin32 and re
# Python script to get and parse Fortigate's IPS email alerts stored in a specific folder inside inbox.
# This script is looking for a unique source ip address in a body and get all related to him attacks and their dates&times.
# This script also displays the source country of the attacker ip address and the sassionid in forigate logs.
# Wrote by J.S 
# Enjoy

import win32com.client
import re


def extract_email_data():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the desired folder in inbox:
    inbox = outlook.GetDefaultFolder(6).Folders("Blacklist")

    # Get all the items in the folder:
    mails = inbox.Items

    email_data = {}

    # Iterate over each mail in the folder
    for mail in mails:
        if mail.Class == 43:  # Check  if the mail is a mail item
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

# Test the function
emails = extract_email_data()

# Print the extracted email data
for srcip, data in emails.items():
    print(f"Source IP: {srcip}")
    print(f"Dates: {data['date']}")
    print(f"Times: {data['time']}")
    print(f"Source Countries: {data['srccountry']}")
    print(f"Session IDs: {data['sessionid']}")
    print(f"Attack Method: {data['attack']}")
    print(f"Destination IPs: {data['dstip']}")
    print()
