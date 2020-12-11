# -*- coding: utf-8 -*-
"""
Created on Mon Aug 27 11:43:26 2018

@author: Sarah Zeng

Automatically send pair-wise emails to each COMRADES pairing
Purpose: Notification of pairing
"""

import pandas as pd
import smtplib
import getpass
from email.message import EmailMessage

# Email credentials
me = input('Enter your email address:\n')
pwd = getpass.getpass(prompt='Enter your password:\n')

# Login to server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(me, pwd)

# Import pair and contact information
pairsheet = 'pairs.xlsx'
df = pd.read_excel(pairsheet)
pairCount = df['Pair'].max()

# Housecleaning
df.Year.where(df.Year==1, 0, inplace=True)
df = df.sort_values(['Pair', 'Year']).set_index(['Pair', 'Year'])
name = df['First Name'] + ' ' + df['Last Name']
df.insert(0, 'Name', name)
df.drop(['First Name', 'Last Name'], inplace=True, axis=1)

def compose_contact_info(df, pair):
    """ Creates block of text with given pair's contact information """
    mentor = df.loc[pair, 0].dropna()
    mentee = df.loc[pair, 1].dropna()
    contact_info = mentor.to_string() + '\n\n' + mentee.to_string()
    return contact_info

def sendto(df, pair):
    """ Returns string of email addresses for given pair """
    recipients = df.loc[pair][['UCSD Email', 'Other Email']].stack()
    return ', '.join(list(recipients))

# Load in text for the email body
with open('intro_message.txt', 'r') as file:
    email_text = file.read()

for pair in range(1, pairCount + 1):
    # Preview email
    recipients = sendto(df, pair)
    contact_info = compose_contact_info(df, pair)
    email_body = email_text.format(contact_info, me)
    print(email_body)
    
    # Compose email
    msg = EmailMessage()
    msg['Subject'] = 'You have a COMRADE!'
    msg['From'] = me
    msg['To'] = recipients
    msg.set_content(email_body)
    
    # Verification
    send_text = f'Would you like to send this email to [{recipients}]? (Y/N/stop)\n'
    send_email = input(send_text)
    
    if send_email == 'Y':
        send_email = input('Are you sure you would like to send this message?\n')
        if send_email == 'Y':
            server.send_message(msg)
            print('\n\n Email sent! \n\n\n')
        else:
            break
    elif send_email == 'stop':
        break
    else:
        print('\n\n Email not sent.\n\n')
        
print('\nCode complete.')