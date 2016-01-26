#!/usr/bin/env python3

import re
import io
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import PyPDF2


'''Remember to enable access for "less secure devices" in your gmail settings!
'''

TIMETABLES_PDF = ''
EMAIL_USER = ''
EMAIL_PASS = ''


def get_email_address(oen, last_name, first_name):
    return '{0}{1}{2}@ugcloud.ca'.format(first_name[:2],
                                         last_name[:3],
                                         oen.replace('-','')[5:])


def send_email(email_address, pdf_io, smtp):
    msg = MIMEMultipart()
    msg['Subject'] = 'Your Timetable for Semester 2'
    msg['To'] = email_address
    pdf_io.seek(0)
    msg.attach(MIMEApplication(pdf_io.read(), _subtype='PDF'))
    smtp.send_message(msg)


def main(pdf_file=TIMETABLES_PDF):
    reader = PyPDF2.PdfFileReader(open(pdf_file, 'rb'))
    oen_re = re.compile('OEN Number: .*[0-9]{3}-[0-9]{3}-[0-9]{3}')
    smtp = SMTP_SSL('smtp.gmail.com:465')
    smtp.login(EMAIL_USER, EMAIL_PASS)
    for page in reader.pages:
        text = page.extractText()
        match = oen_re.search(text)
        if match:
            oen = text[match.start()+11:match.end()].strip()
            full_name = text[35:match.start()].strip()
            if not full_name.count(',') == 1:
                print('{0} has a weird name'.format(full_name))
                continue
            else:
                last_name, first_name = full_name.split(',')
            email_address = get_email_address(oen,
                                              last_name.strip(),
                                              first_name.strip())
            writer = PyPDF2.PdfFileWriter()
            writer.addPage(page)
            pdf_io = io.BytesIO()
            writer.write(pdf_io)
            send_email(email_address, pdf_io, smtp)
    smtp.close()


if __name__ == '__main__':
    main()
