import email
import imaplib
import os
import shutil
import sqlite3
from datetime import datetime

import pdfkit

from make_log import log_exceptions

db, folder = 'database1.db', 'temp/'


def process_values(fromtime, totime, insname):
    """
    1.Accept values
    2.Get table of insname, hospital, subject, attachment path
    3.Check if file exists in attachment path
    4.If exists then call make_excel
    5.If not exists then call check_and_download_attachment
    6.Download attachment and call make_excel
    7.if excel created then call process_insurer_pdfs
    :param fromtime:
    :param totime:
    :param insname:
    :return:
    """
    try:
        record = []
        if insname != 'all':
            q1 = f"select row_no,insurerid,hos_id,date,emailsubject,file_path  from updation_detail_log" \
                 f" where date between '{fromtime}' and '{totime}'" \
                 f" and fieldreadflag='X' and insurerid='{insname}' order by row_no"
        else:
            q1 = f"select row_no,insurerid,hos_id,date,emailsubject,file_path  from updation_detail_log" \
                 f" where date between '{fromtime}' and '{totime}'" \
                 f" and fieldreadflag='X' order by row_no"

        with sqlite3.connect("database1.db") as con:
            cur = con.cursor()
            cur.execute(q1)
            r = cur.fetchall()
            if r is not None:
                for row in r:
                    record.append(check_and_download_attachment(str(row[0]), row[1], row[2], row[4]))
                    for i in record:
                        with open("records.csv", "a+") as fp:
                            i = str(i).replace("(", "").replace(")", "")
                            fp.write(i+'\n')
            pass
    except:
        log_exceptions()
        pass


def check_and_download_attachment(row_no, insname, hospital, subject):
    flag = 0
    try:
        shutil.rmtree(folder, ignore_errors=True)
        os.mkdir(folder)
        dst_directory = 'backups/'
        date_time = datetime.now().strftime("%m%d%Y%H%M%S")
        finaldirectory = dst_directory + insname + '_' + date_time
        if not os.path.exists(dst_directory):
            os.mkdir(dst_directory)
        if not os.path.exists(finaldirectory):
            os.mkdir(finaldirectory)
        server, email_id, password, inbox = "", "", "", ""
        if 'Max' in hospital:
            server, email_id, password, inbox = "outlook.office365.com", "Tpappg@maxhealthcare.com", "Sept@2020", '"Deleted Items"'
        elif 'inamdar' in hospital:
            server, email_id, password, inbox = "imap.gmail.com", "mediclaim@inamdarhospital.org", "Mediclaim@2019", '"[Gmail]/Trash"'
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_id, password)
        mail.select('inbox', readonly=True)
        subject = subject.replace("\r", "").replace("\n", "")
        type, data = mail.search(None, f'(SUBJECT "{subject}")')
        if data == [b'']:
            mail.select(inbox, readonly=True)
            type, data = mail.search(None, f'(SUBJECT "{subject}")')
        mid_list = data[0].split()
        if mid_list == []:
            return row_no, flag, insname, hospital, subject
        result, data = mail.fetch(mid_list[-1], "(RFC822)")
        raw_email = data[0][1].decode('utf-8')
        email_message = email.message_from_string(raw_email)
        subject = email_message['Subject']
        for mail.part in email_message.walk():
            if mail.part.get_content_type() == "text/html" or mail.part.get_content_type() == "text/plain":
                mail.body = mail.part.get_payload(decode=True)
                mail.file_name = folder + 'email.html'
                mail.output_file = open(mail.file_name, 'w')
                mail.output_file.write("Body: %s" % (mail.body.decode('utf-8')))
                mail.output_file.close()
                pdfkit.from_file(folder + 'email.html', folder + row_no + '.pdf')
                if os.path.exists(folder + 'email.html'):
                    os.remove(folder + 'email.html')
            filename = mail.part.get_filename()
            if filename is not None:
                if os.path.exists(folder + row_no + '.pdf'):
                    os.remove(folder + row_no + '.pdf')
                if os.path.exists(folder + 'email.html'):
                    os.remove(folder + 'email.html')
                att_path = os.path.join(folder, filename)
                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    fp.write(mail.part.get_payload(decode=True))
                    fp.close()
        pass
        file_names = os.listdir(folder)
        for file_name in file_names:
            flag += 1
            shutil.move(os.path.join(folder, file_name), finaldirectory)
        return row_no, flag, insname, hospital, subject
    except:
        log_exceptions(subject=subject)
        return row_no, flag, insname, hospital, subject


if __name__ == "__main__":
    process_values('01/08/2020', '18/09/2020', 'all')
    # a = check_and_download_attachment('12', 'Raksha', 'Max',
    #                               "Claim Settlement Letter From Raksha Health Insurance TPA Pvt.Ltd. (M58ADD676ILBS,9331198,Dr. Aditi Goel.)")
    pass
