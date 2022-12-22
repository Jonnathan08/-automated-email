# %%
#!pip install sharepy
import sharepy
import os
#System variables
MAIL_NUB = os.environ.get('MAIL_NUB')
PASS_MAIL_NUB = os.environ.get('PASS_MAIL_NUB')


s = sharepy.connect("sharepoint directory", MAIL_NUB, PASS_MAIL_NUB)

# %%
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from pandas import Timestamp
import numpy as np
import sharepy
import create_image as img

from email import encoders
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

# %%
#Importing data from CA tracker
r = s.getfile("sharepoint directory file"\
              , filename = 'ca_tracker.xlsx')
df = pd.read_excel("ca_tracker.xlsx", usecols=['Name' , 'Start Date' , 'Finish Date', 'Out of Shift', 'Future Dates', 'Effort Calculation', 'Blank values'], sheet_name='CA Tracker')
df.dropna(subset = ['Start Date'], inplace = True)
df.dropna(subset = ['Name'], inplace = True)

# %%
#import data with the users mails
p = s.getfile("sharepoint directory file"\
              , filename = 'users_mails.xlsx')
user_mails = pd.read_excel ('users_mails.xlsx')
print('done')

# %%
#SMTP Mail variables

myMail = 'username'
password = 'password'
subjectEmail = 'CA Tracker Reminder'

msg1 = """<h1>Hola, %s</h1>
        <h3>A continuaci&oacute;n un resumen de tu actividad en CA_Tracker</h3>
        <p><img src="imagedraw.png" /></p>
        <h3>Por favor revisa el archivo adjunto para encontrar los errores</h3>
        <h3>Asegurate de que tus registros esten corregidos y al d&iacute;a</h3>"""
        
msg2 = """<h1>Hola, %s</h1>
        <h3>A continuaci&oacute;n un resumen de tu actividad en CA_Tracker</h3>
        <p><img src="imagedraw.png" /></p>
        <h3>Por favor revisa el archivo adjunto para encontrar los errores</h3>
        <h3>Asegurate de corregir las l&iacute;neas con errores</h3>"""
        
msg3 = """<h1>Hola, %s</h1>
        <h3>A continuaci&oacute;n un resumen de tu actividad en CA_Tracker</h3>
        <p><img src="imagedraw.png" /></p>
        <h3>Asegurate de mantener tus registros al d&iacute;a</h3>"""
        
msg4 = """<h1>Hola, %s</h1>
        <h3>A continuaci&oacute;n un resumen de tu actividad en CA_Tracker</h3>
        <p><img src="imagedraw.png" /></p>
        <h3>¡Gracias por mantener tu información al día!</h3>"""

# %%
#Get the total of names in the CA tracker file and elimnate duplicates

total_names = df["Name"].drop_duplicates().to_frame()
#total_names.to_excel("total_names.xlsx")
print('done')

# %%
#Formating Start Date column in CA tracker
df['Start Date']=df['Start Date'].apply(lambda x: datetime.strptime(str(x) , '%Y-%m-%d %H:%M:%S')  if type(x) == datetime or type(x) == Timestamp else datetime.strptime('1900-01-01 00:00:00' , '%Y-%m-%d %H:%M:%S'))
print('done')

# %%
#Getting the variables of yerterday

#la respuesta se recibe como un entero en donde 0 es Lunes y 6 es Domingo (week_day)
week_day = datetime.today().weekday()

if week_day == 0:
    yesterday = datetime.today() - timedelta(days=3)
else:
    yesterday = datetime.today() - timedelta(days=1)
    
print(yesterday.day)    
today = datetime.today()
date = str(today.year)+'-'+str(today.month)+'-'+str(today.day)

# %%
#filter the users that have registers yesterday

no_mail = pd.DataFrame()

for i, r in df.iterrows():
    if (r[1].year == yesterday.year) and (r[1].month == yesterday.month) and (r[1].day == yesterday.day):
        no_mail = no_mail.append(r) 
print('done')

# %%
if len(no_mail) > 0:
    #Filtring names of the users (with register yesterday) and eliminating duplicates
    no_mail_names = no_mail["Name"].drop_duplicates().to_frame()

    #convining tables to get the names of the users that do not have registers yesterday
    s_mail_names = pd.merge(total_names, no_mail_names, on=['Name', 'Name'], how='outer', indicator=True).query('_merge=="left_only"')
    sent_mail_names = s_mail_names['Name'].to_frame()

    #adding new column 'not updated'
    sent_mail_names['Not updated'] = 1
else:
    sent_mail_names = total_names['Name'].to_frame()
    sent_mail_names['Not updated'] = 1

print('done')

# %%
##Last date register

# %%
#Getting the last date register in the ca tracker per user
last_reg = df[['Name','Start Date']].groupby('Name').max()

# %%
##Errors

# %%
#Search for user with error in their registres
users_errors = pd.DataFrame()

for i, s in df.iterrows():
    if (s[3] == 'Error' or s[4] == 'Error' or s[5] == 'Error' or s[6] == 'Error'):
       users_errors = users_errors.append(s)
       
print('done')

# %%
#Counting the number of rows with errors, named 'Line with errors'
line_errors = users_errors.groupby('Name').agg({'Start Date': 'count'}).reset_index().rename(columns = {'Start Date': 'Line with error'})

# %%
#Mapping specific errors
users_errors['Error row'] = users_errors.index
users_errors['Error row'] = users_errors['Error row'] + 2
users_errors.set_index('Error row',inplace=True)
users_errors
users_errors[['Name','Out of Shift', 'Future Dates','Effort Calculation', 'Blank values']].to_excel('ca_tracker_errors_%s.xlsx' % date)

# %%
##Merging data
final_table = total_names.merge(line_errors, how = 'left', left_on = 'Name', right_on= 'Name').merge(
                sent_mail_names, how = 'left', left_on = 'Name', right_on= 'Name').merge(
                user_mails, how = 'left', left_on = 'Name', right_on= 'Name').merge(
                last_reg, how = 'left', left_on = 'Name', right_on= 'Name')

# %%
final_table.to_excel("names_to_send_mail.xlsx")

# %%
final_table

# %%
##Starting outlook server connection

# %%
#outlook server connection
outlServer = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)

#Initialize encryption
outlServer.starttls()

#Login in the outlook account using the define variables username and password
outlServer.login(myMail, password)

# %%
#Definition of the function to send emails
def send_msg(To, From, Subject, Message, s_attach):
    msg = MIMEMultipart()
    msg['From'] = From
    msg['To'] = To
    msg['Subject'] = Subject
    msg_content = MIMEText(Message, 'html', 'utf-8')
    msg.attach(msg_content)

    #attach image // the image is save in the same directory
    with open('imagedraw.png', 'rb') as f:
        mime = MIMEBase('image', 'png', filename='imagedraw.png')
        mime.add_header('content-Disposition', 'attachment', filename='imagedraw.png')
        mime.add_header('X-Attachment-ID', '0')
        mime.add_header('ContextID', '<0>')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        msg.attach(mime)
        
    if s_attach == True:
        with open('ca_tracker_errors_%s.xlsx' % date, 'rb') as e:
            exc = MIMEBase('application', 'octet-stream', filename='ca_tracker_errors_%s.xlsx' % date)
            exc.add_header('content-Disposition', 'attachment', filename='ca_tracker_errors_%s.xlsx' % date)
            exc.set_payload(e.read())
            encoders.encode_base64(exc)
            msg.attach(exc)
    
    outlServer.sendmail(From, To, msg.as_string())

def format_time(date):
    format = str(date.month)+'/'+str(date.day)+'/'+str(date.year)
    return format

# %%
#Sending mail according with establish cases

mails = pd.DataFrame()

for i, m in final_table.iterrows():
    #Have two conditions
    if(pd.isnull(m[1]) == False and m[2] == 1):
        #Validate if there is a mail address
        if(pd.isnull(m[3]) == False):
            print(m[0] + ' con 1')
            img.generate_image(m[0], format_time(m[5]) ,int(m[1]))
            send_msg(m[3], myMail, subjectEmail ,msg1 % m[0], True)
    #Have only errors
    elif (pd.isnull(m[1]) == False and m[2] != 1):
        if(pd.isnull(m[3]) == False):
            print(m[0] + ' con 2')
            img.generate_image(m[0], format_time(m[5]) ,int(m[1]))
            send_msg(m[3], myMail, subjectEmail ,msg2 % m[0], True)
    #Not updated yesterday
    elif (pd.isnull(m[1]) == True and m[2] == 1):
        if(pd.isnull(m[3]) == False):
            print(m[0] + ' con 3')
            img.generate_image(m[0], format_time(m[5]) , 0)
            send_msg(m[3], myMail, subjectEmail ,msg3 % m[0], False)
    elif (pd.isnull(m[1]) == True and pd.isnull(m[2]) == True):
        if(pd.isnull(m[3]) == False):
            print(m[0] + ' con 4')
            img.generate_image_congrat()
            send_msg(m[3], myMail, subjectEmail ,msg4 % m[0], False)
            
print('done')

# %%
outlServer.quit()

# %%



