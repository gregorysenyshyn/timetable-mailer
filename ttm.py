import re
import io
import time
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

import PyPDF2

PDF_FILE = ''
EMAIL_USER = ''
EMAIL_PASS = ''
EMAIL_SUBJECT = ''
EMAIL_BODY = ''

def strip_punctuation(full_name):
    things_to_remove = ["'", " ", "-"]
    for thing in things_to_remove:
        full_name.replace(thing, "")
    return full_name


def get_email_address(oen, full_name):
    full_name = strip_punctuation(full_name)
    last_name, first_name = full_name.split(',')
    return '{0}{1}{2}@ugcloud.ca'.format(first_name.strip().capitalize()[:2],
                                         last_name.strip().capitalize()[:3],
                                         oen.replace('-', '')[5:])


def get_oen_re():
    return re.compile('OEN Number: .*[0-9]{3}-[0-9]{3}-[0-9]{3}')


def get_smtp(user, password):
    smtp = SMTP_SSL('smtp.gmail.com:465')
    smtp.login(user, password)
    return smtp


def get_message(page, email_user, subject, oen_re, body=None):
    text = page.extractText()
    match = oen_re.search(text)
    if match:

        writer = PyPDF2.PdfFileWriter()
        writer.addPage(page)
        pdf_io = io.BytesIO()
        writer.write(pdf_io)
        pdf_io.seek(0)

        oen = text[match.start()+11:match.end()].strip()
        full_name = text[35:match.start()].strip()
        full_name = strip_punctuation(full_name)
        email_address = get_email_address(oen, full_name)

        msg = MIMEMultipart()
        msg['To'] = email_address
        msg['From'] = email_user
        msg['Subject'] = subject
        if body is not None:
            msg.attach(MIMEText(body))
        msg.attach(MIMEApplication(pdf_io.read(), _subtype='PDF'))

        return msg


def main():
    with open(PDF_FILE, 'rb') as f:
        reader = PyPDF2.PdfFileReader(f)
        oen_re = get_oen_re()
        smtp = get_smtp(USER_EMAIL, USER_PASS)
        for page in reader.pages:
            msg = get_message(page, EMAIL_USER,
                              EMAIL_SUBJECT, oen_re, EMAIL_BODY)
            smtp.send_message(msg)
            print('Sent email to: '.format(msg['To']))
            time.sleep(2)
        smtp.close()


if __name__ == '__main__':
    main()
