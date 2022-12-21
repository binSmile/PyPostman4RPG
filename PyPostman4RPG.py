import os

import gspread, datetime
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
import requests
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email.headerregistry import Address
from email import encoders
from PyPDF2 import PdfFileReader, PdfFileWriter


def main():
    scope = ['https://spreadsheets.google.com/feeds']
    SENDER = 'SENDER MAIL ADRESS'
    CONTENT = 'MESSAGE CONTENT'
    PATH_TO_JSON_KEYFILE = 'filename.json'
    sheetURL = 'https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXX/edit'
    TMP_PATH = '/tmp/'
    SMTP_PASSWORD = 'XXXXXXXXX'
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 587
    TOPIC_PREFIX = '[MyRPG]'

    creds = ServiceAccountCredentials.from_json_keyfile_name(PATH_TO_JSON_KEYFILE, scope)
    gc = gspread.authorize(creds)
    sht2 = gc.open_by_url(sheetURL)
    worksheet = sht2.get_worksheet(0)
    list_of_lists = worksheet.get_all_values()

    adrBook = sht2.get_worksheet(1)
    list_of_contacts = adrBook.get_all_values()

    service = build('drive', 'v3', credentials=creds)
    mailCounter = 0

    i = 1
    for mail in list_of_lists[1:]:
        i += 1
        if mail[5] and not mail[7]: # real email and not date of send
            file_id = mail[4].split('=')[1]
            url = 'https://drive.google.com/uc?export=download&id=%s' % file_id
            response = requests.get(url)
            header = response.headers['Content-Disposition']
            file_name = re.search(r'filename="(.*)"', header).group(1)
            with open(TMP_PATH+file_name, 'wb') as f:
                f.write(response.content)

            if file_name.split('.')[-1] == 'pdf':
                try:
                    fin = open(TMP_PATH+file_name, 'rb')
                    reader = PdfFileReader(fin)
                    writer = PdfFileWriter()
                    writer.appendPagesFromReader(reader)
                    metadata = {'/Author': 'RPG postman'}
                    writer.addMetadata(metadata)
                    fout = open(TMP_PATH+'a_'+file_name, 'wb')
                    writer.write(fout)

                    fin.close()
                    fout.close()
                    os.replace(TMP_PATH+'a_'+file_name,TMP_PATH + file_name)

                except:
                    worksheet.update_cell(i, 9, 'Error with PDF anonimizer >_<')

            sendQ = False
            for contact in list_of_contacts:
                if contact[0] == mail[5]:
                    sendQ = True
                    try:
                        worksheet.update_cell(i, 7, contact[1])
                        send_message(SENDER, contact[1], TOPIC_PREFIX + ' '+ + mail[3] , CONTENT, TMP_PATH+ file_name)
                        mailCounter += 1
                        if mail[7] == 'email was not found':
                            worksheet.update_cell(i, 9, '')
                    except:
                        worksheet.update_cell(i, 9, 'Sent error >_<')
            if not sendQ:
                worksheet.update_cell(i, 9, 'email was not found')
            os.remove(TMP_PATH+ file_name)
            worksheet.update_cell(i,8, str(datetime.datetime.now()))
    worksheet.update_cell(1, 11, str(datetime.datetime.now()))
    print('Carrier pigeon flew away with %d letters.' % mailCounter)


def send_message(MY_EMAIL,TO_EMAIL,SUBJECT,message_text,FILEPATH):
    msg = MIMEMultipart()
    # msg['From'] = Address("RPG Mail", MY_EMAIL)
    msg['From'] = MY_EMAIL
    msg['To'] = COMMASPACE.join([TO_EMAIL])
    msg['Subject'] = SUBJECT
    messaga = MIMEText(message_text, 'plain', 'utf-8')
    msg.attach(messaga)


    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(FILEPATH, "rb").read())
    encoders.encode_base64(part)
    filename_true = os.path.basename(FILEPATH)
    filename = 'letter.' + filename_true.split('.')[-1]
    part.add_header('Content-Disposition', 'attachment', filename=filename)  # or
    msg.attach(part)

    smtpObj = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(MY_EMAIL, SMTP_PASSWORD)
    smtpObj.sendmail(MY_EMAIL, TO_EMAIL, msg.as_string())
    smtpObj.quit()




main()