import pandas as pd
import smtplib
import ssl
from email.message import EmailMessage


#PLEASE READ
#This program reads property adresses and emails from an Excel file, and 
#sends an email to each person. This program requires a gmail account with 2-FA enabled to work.
#After enabling 2-FA on your google account, create a custom device on the same page 
#and use the password that google generates as the password in this script.


pd.read_excel('add-excel-file-here')


#Put the column containing the property address first, and the column containing the Email 2nd.
df = pd.read_excel('add-excel-file-here', usecols=[0,9])


data = [] 

email_sender = 'add-your-email-here'
email_password = 'add-your-password-here'

for val in d.values:
    data.append(val)
    
for d in data:
    address = d[0]
    email_recipient = d[1]
    subject = 'add-subject-here'
    body = "hello are you interested in selling your house on " + address + " thank you so much "
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_recipient
    em['Subject'] = subject
    em.set_content(body)

    #Adding SSL
    context = ssl.create_default_context()

    #Sending Email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_recipient, em.as_string())
