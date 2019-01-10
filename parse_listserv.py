#!/usr/bin/env python3
# Adapted from https://github.com/jdcc/parse_listserv

"""Unpack a MIME message into a directory of files."""

import os
import sys
import email
import errno
import mimetypes
import csv
import re
import datetime

from argparse import ArgumentParser

EMAIL_SEP = '========================================================================='
DATE_FORMAT = '%a, %d %b %Y %H:%M:%S %z'
DATE_FORMAT_2 = '%a, %d %b %Y %H:%M:%S %Z'
OUT_DATE_FORMAT = '%Y-%m-%d %H:%M:%S%z'
CSV_HEADERS = ['Date', 'Subject', 'From', 'Directory', 'Attachments']

def make_one_line(txt):
    if txt:
        return re.sub('\s+', ' ', re.sub('\\r|\\n', '', txt))
    else:
        return ''

def dump_email(txt, folder, csvwriter):
    try:
        os.makedirs(folder)
    except FileExistsError:
        pass

    msg = email.message_from_string(txt.encode("ascii", errors="ignore").decode())

    counter = 1
    filenames = []
    for part in msg.walk():
        # multipart/* are just containers
        if part.get_content_maintype() == 'multipart':
            continue
        # Applications should really sanitize the given filename so that an
        # email message can't be used to overwrite important files
        filename = part.get_filename()
        if filename:
            filenames.append(filename)
        else:
            ext = mimetypes.guess_extension(part.get_content_type())
            if not ext:
                # Use a generic bag-of-bits extension
                ext = '.bin'
            filename = 'part-%03d%s' % (counter, ext)
        counter += 1
        with open(os.path.join(folder, filename), 'wb') as fp:
            try:
                fp.write(part.get_payload(decode=True))
            except TypeError as e:
                print(msg['Subject'], '-', e)
            except UnicodeError as e:
                print(msg['Subject'], '-', e)

    try:
        date = datetime.datetime.strftime(
                datetime.datetime.strptime(make_one_line(msg['Date']), DATE_FORMAT),
                OUT_DATE_FORMAT)
    except:
        date = datetime.datetime.strftime(
                datetime.datetime.strptime(make_one_line(msg['Date']), DATE_FORMAT_2),
                OUT_DATE_FORMAT)

    csvwriter.writerow([
        date,
        make_one_line(msg['Subject']),
        make_one_line(msg['From']),
        str(folder),
        "; ".join(filenames)
        ])

def parse_logfile(logfile, out, csvwriter):
    try:
        os.makedirs(out)
    except FileExistsError:
        pass

    email = ""
    counter = 1
    for line in open(logfile, newline='', errors="surrogateescape"):
        if line.strip() == EMAIL_SEP:
            if len(email) > 0:
                try:
                    dump_email(email, os.path.join(out, str(counter)), csvwriter)
                    counter += 1
                    email = ""
                except:
                    pass
        else:
            email += line

def main():
    parser = ArgumentParser(description=
            "Unpack a listserv .log directory into emails")
    parser.add_argument('logdir',
                        help="""Directory containing listserv .log files""")
    parser.add_argument('outdir',
                        help="""Directory to contain emails""")
    args = parser.parse_args()

    try:
        os.makedirs(args.outdir)
    except FileExistsError:
        pass

    csvfile = open(os.path.join(args.outdir, 'emails.csv'), 'w')
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(CSV_HEADERS)

    for logfile in os.scandir(args.logdir):
        if not logfile.is_file():
            continue
        logfiledir = os.path.join(args.outdir, logfile.name.replace('.log', ''))
        parse_logfile(os.path.join(args.logdir, logfile.name), logfiledir, csvwriter)

    csvfile.close()

if __name__ == '__main__':
    main()
