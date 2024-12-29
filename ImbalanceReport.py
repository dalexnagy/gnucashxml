#!/usr/bin/python3
#----------------------------------------------------------------------------------------------
# Report if Imbalance-USD account has any transactions in it
# Intended to run as a daily task and send email to me if any transactions are found
#
# --- Change History ---
# Program version 0
# 2024-12-24 V1   - New


Program_Version = "V1.0"

# System imports
import sys
from datetime import date
import smtplib
import ssl
# GnuCash Structure import
import gnucashxml

sys.path.insert(0, '/home/dave/Python/_Configs')
import _EMail

# CONSTANTS & Globals
account_to_check = "Imbalance-USD"
# Get current date and set other variables
today = date.today()
notify_email = ["dalexnagy@gmail.com"]
# END OF CONSTANTS

# -------------

book = gnucashxml.from_filename("/home/dave/GnuCash/NagyFamily2024.gnucash")


for account, children, splits in book.walk():
    if account.name == account_to_check:
        if len(splits) > 0:
            notify_msg = "Warning: {} Transactions were found in '{}' account".format(len(splits), account_to_check)
        else:
            notify_msg = "No Transactions Found in '{}' account".format(account_to_check)


# Setup Email connection and message
port = 587  # For starttls
smtp_server = "smtp.gmail.com"
message = """Subject: '{}' Status

""".format(account_to_check)
context = ssl._create_unverified_context()
# Send notification email
message = message + notify_msg
with smtplib.SMTP(smtp_server, port) as server:
    server.starttls(context=context)
    server.login(_EMail.email_user, _EMail.email_password)
    server.sendmail(_EMail.email_user, notify_email, message)