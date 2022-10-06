#!/usr/bin/env python3

import openpyxl, smtplib, sys, datetime

# Open the spreadsheet and get the latest dues status.
file = 'garden_TEST.xlsx'
wb = openpyxl.load_workbook(file)
ws = wb.active

# Set today's date
ds = datetime.datetime.now().strftime('%Y-%m-%d') # Today's date

# gmail app-specific password
username = 'USERNAME GOES HERE'
pw = 'PASSWORD GOES HERE'

# Iterate through column C(third column), if there is an X store name and email values
for row in ws.iter_rows(min_row=1, max_col=3, max_row=11, values_only=True):
    for cell in row:
        if cell == 'x':
            name = row[0]
            email = row[1]
            print()
            print("############################")
            print(ds)
            print("Today's waterer was " + name)
            print("Email sent to " + email)
            print("############################")
            print()

#Log in to email account
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login(username, pw ) # password is a gmail app-specific password

# send email, close email connection
body = "Subject: It's your turn to water the garden!\nHello %s,\n\nToday is your day to water MHL's garden. \n" \
        "\nIf you can't water it today, please find someone else who can. \n\nIf there was a decent amount of rainfall " \
        "in the last 24 hours and things seem damp, you probably don't need to water today. \n\nThank you!" % (name)
# print('Sending email to %s...' % email)
sendmailStatus = smtpObj.sendmail('mhlgardengnome@gmail.com', email, body)
smtpObj.quit()


# delete column C and save new version of file
ws.delete_cols(3)
wb.save('garden_TEST.xlsx')








