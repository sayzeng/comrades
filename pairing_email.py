# -*- coding: utf-8 -*-
"""
Created on Mon Aug 27 11:43:26 2018

@author: Sarah Zeng

Automatically send pair-wise emails to each COMRADES pairing
Purpose: Notification of pairing
"""

import pandas as pd
import numpy as np
import smtplib
from email.message import EmailMessage

#email credentials
me=input('Enter your email address:')
pwd=input('Enter your password:')
#definitions
pairsheet='pairs.xlsx'

#login to server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(me,pwd)

#import pair and contact information
df0=pd.read_excel(pairsheet)
pairs=df0['Pair'].max()

#misc
com=', '
tab='\t'
nsep='\n'
nnsep='\n \n'

#seed
df=pd.DataFrame()
columns=['PairID','contactinfo','sendto']

#START HERE #spit out: string of to emails (you), table for email body (line2)
for i in range(1,pairs+1):
    pairinfo=df0.loc[df0['Pair']==i].replace(np.nan,'N/A')
    
    n='Name: '
    n1=pairinfo.iloc[0]['First Name']+' '+pairinfo.iloc[0]['Last Name']
    n2=pairinfo.iloc[1]['First Name']+' '+pairinfo.iloc[1]['Last Name']
    p='Phone: '
    p1=pairinfo.iloc[0]['Phone Number']
    p2=pairinfo.iloc[1]['Phone Number']
    e='Email: '
    e1a=pairinfo.iloc[0]['UCSD Email']
    e1b=pairinfo.iloc[0]['Other Email']
    e2a=pairinfo.iloc[1]['UCSD Email']
    e2b=pairinfo.iloc[1]['Other Email']
    e1=e1a+com+e1b
    e2=e2a+com+e2b
    
    c1=n+tab+n1+nsep +p+tab+p1+nsep +e+tab+e1
    c2=n+tab+n2+nsep +p+tab+p2+nsep +e+tab+e2
    
    no=pairinfo.iloc[0]['Pair']
    c=c1+nsep+nsep+c2
    sendto=e1+com+e2
    
    row=pd.DataFrame([no,c,sendto]).transpose()
    df=df.append(row)

df.columns=columns  
df=df.reset_index(drop=True)
print('Ready to construct emails.\n')

#iteration
for i in range(pairs):
    #email formatting
    msg = EmailMessage()
    msg['Subject']='You have a COMRADE!'
    msg['From']=me
    
    #compose email body
    line1='Welcome to the UCSD Economics COMRADES mentorship program! I hope that you both find your mentorship experience rewarding and beneficial. You can find each other’s contact information below:'
    line3='For the month of September, you each have $5 ($10 total) to spend on coffee or froyo or any other kind of bonding consumption good (except alcohol, sorry!). To submit a reimbursement, please take a photo of your itemised receipt and send it to '+me+' with the string ‘COMRADES’ in the subject line.'
    line4='\n Sarah'
    
    #iterated content
    msg['To']=df.iloc[i]['sendto']
    line2=df.iloc[i]['contactinfo']
    lines=(line1,line2,line3,line4)
    body=nnsep.join(lines)
    msg.set_content(body)
    print(body)
    
    #Warning input
    warningtxt='Are you sure you want to send an email to '+df.iloc[i]['sendto']+' ? (Y/N)\n'
    send_email=input(warningtxt)
    
    if send_email=='Y':
        server.send_message(msg)
        print('\n\n Email sent! \n\n\n')
    elif send_email=='break':
        break
    else:
        print('\n\n Check your code again.')