import os
import re
import io
import sys
import argparse

import PyPDF2


def strip_punctuation(full_name):
    things_to_remove = ["'", "-", ".", " "]
    for thing in things_to_remove:
        full_name = full_name.replace(thing, "")
    return full_name


def get_oen_re():
    return re.compile('OEN Number: +[0-9]{3}-[0-9]{3}-[0-9]{3}')


def get_ugcloud_re():
    return re.compile('^[a-zA-Z]{5}[0-9]{4}$')


def get_username(text, oen_match):
    oen = text[oen_match.start()+11:oen_match.end()].strip()
    full_name = strip_punctuation(text[35:oen_match.start()].strip())
    last_name, first_name = full_name.split(',')
    return '{0}{1}{2}'.format(first_name.strip().capitalize()[:2],
                              last_name.strip().capitalize()[:3],
                              oen.replace('-', '')[5:])

def close_document(writer, output_dir, current_ugcloud):
    outfile = f"{current_ugcloud}.pdf"
    outpath = os.path.join(output_dir, outfile)
    with open(outpath, "wb") as f:
        writer.write(f)
        print(f"  Done!")
    return PyPDF2.PdfFileWriter()


def file_checker(args, value_type, message):
    if vars(args)[value_type] is not None: 
        if os.path.exists(vars(args)[value_type]):
            return vars(args)[value_type]
    else:
        value = None
        while not value:
            value = input(f"{message} location: ").strip()
            if os.path.exists(value):
                return value
            else:
                print(f"{value} is not a valid location")
                value = None

def main(filename, output_dir):

    oen_re = get_oen_re()
    ugcloud_re = get_ugcloud_re()

    reader = None
    with open(filename, 'rb') as f:
        reader = PyPDF2.PdfFileReader(f)

        writer = PyPDF2.PdfFileWriter()
        current_ugcloud = None
        new_document = True

        print("processing pages...\n")
        for page in reader.pages:
            text = page.extractText()
            oen = oen_re.search(text)
            if oen:
                email_address = get_username(text, oen)
                ugcloud = ugcloud_re.search(email_address)
                print(f"processing {email_address}...", end="")

                if not ugcloud:
                    print(f"{email_address} is not a valid UGCloud address")
                    
                if new_document:
                    current_ugcloud = email_address
                else:
                    writer = close_document(writer,
                                            output_dir,
                                            current_ugcloud)
                    new_document = True
                    current_ugcloud = email_address
                writer.addPage(page)
                new_document = False

            else:
                writer.addPage(page)
        close_document(writer, output_dir, current_ugcloud)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=("Take a Maplewood timetable "
        "pdf and split it into individual files named with a student's ugcloud"
        " user name"))
    parser.add_argument("-f", "--filename", dest="filename")
    parser.add_argument("-o", "--output", dest="output")
    args = parser.parse_args()
    filename = file_checker(args, "filename", "Maplewood Timetable")
    output = file_checker(args, "output", "Output")

    main(filename, output)
