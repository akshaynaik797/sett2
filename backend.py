import email
import imaplib
import os
import shutil
import sqlite3
import re
import subprocess
from datetime import datetime

import pdfkit

from make_log import log_exceptions
from movemaster import move_master_to_master_insurer


db, folder, dst_directory, directory = 'database1.db', 'temp/', 'backups/', 'backups'

inslist = ('all', 'aditya', 'apollo', 'bajaj', 'big', 'east_west', 'fgh', 'fhpl', 'Good_health', 'hdfc',
           'health_heritage', 'health_india', 'health_insurance', 'icici_lombard', 'MDINDIA', 'Medi_Assist',
           'Medsave', 'Paramount', 'Raksha', 'reliance', 'religare', 'small', 'united', 'Universal_Sompo',
           'vidal', 'vipul')

def process_values(fromtime, totime, insname):
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
                    if os.path.exists(row[5]):
                        date_time = datetime.now().strftime("%m%d%Y%H%M%S")
                        finaldirectory = dst_directory + row[1] + '_' + date_time
                        if not os.path.exists(dst_directory):
                            os.mkdir(dst_directory)
                        if not os.path.exists(finaldirectory):
                            os.mkdir(finaldirectory)
                        shutil.copy(row[5], finaldirectory)
                        temp = row[0], '1', row[1], row[2], row[4]
                        temp = str(temp).replace("(", "").replace(")", "")
                        record.append(temp)
                    else:
                        record.append(check_and_download_attachment(str(row[0]), row[1], row[2], row[4]))
                if os.path.exists("records.csv"):
                    os.remove("records.csv")
                with open("records.csv", "a+") as fp:
                    row = "row no, no of files, insurer, hospital, Email Subject\n"
                    fp.write(row)
                for i in record:
                    with open("records.csv", "a+") as fp:
                        i = str(i).replace("(", "").replace(")", "")
                        fp.write(i + '\n')
        # accept_values(fromtime, datetime.now().strftime("%d/%m/%Y %H:%M:%S"), insname)
        return True
    except:
        log_exceptions()
        return False


def check_and_download_attachment(row_no, insname, hospital, subject):
    flag = 0
    try:
        shutil.rmtree(folder, ignore_errors=True)
        os.mkdir(folder)
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


def accept_values(fromtime, totime, insname):
    fromtime = datetime.strptime(fromtime, '%d/%m/%Y %H:%M:%S')
    totime = datetime.strptime(totime, '%d/%m/%Y %H:%M:%S')
    if insname == 'all':
        for i in inslist:
            if collect_folder_data(fromtime, totime, i):
                print(f'{i} completed')
            else:
                print(f'{i} incomplete')
        return True
    elif collect_folder_data(fromtime, totime, insname):
        return True
    return False


def collect_folder_data(fromtime, totime, insname):
    regex = r'(?P<name>.*(?=_\d+))_(?P<date>\d+)'
    for x in os.walk(directory):
        for y in x[1]:
            if insname in y:
                result = re.compile(regex).search(y)
                if result is not None:
                    tempdict = result.groupdict()
                    folder_insname, foldertime = tempdict['name'], datetime.strptime(tempdict['date'], '%m%d%Y%H%M%S')
                    if fromtime < foldertime < totime and folder_insname == insname:
                        print(f'processing {y}')
                        process_insurer_excel(y, insname, foldertime)
            elif 'star' in y:
                result = re.compile(regex).search(y)
                if result is not None:
                    tempdict = result.groupdict()
                    folder_insname, foldertime = tempdict['name'], datetime.strptime(tempdict['date'], '%m%d%Y%H%M%S')
                    if fromtime < foldertime < totime and folder_insname == 'star':
                        if insname == 'big' or insname == 'small':
                            print(f'processing {y}')
                            process_insurer_excel(y, insname, foldertime)

        break
    return True


def process_insurer_excel(folder_name, insname, foldertime):
    for root, dirs, files in os.walk(directory + '/' + folder_name):
        flag = 0
        for file in files:
            path = (os.path.join(root, file))
            if 'smallinamdar.xlsx' in file:
                op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
                subprocess.run(["python", "make_master.py", 'small_star', op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
            elif 'smallMax.xlsx' in file:
                op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
                subprocess.run(["python", "make_master.py", 'small_star', op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
            elif 'starinamdar.xlsx' in file:
                op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
                subprocess.run(["python", "make_master.py", 'star', op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
            elif 'starMax.xlsx' in file:
                op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
                subprocess.run(["python", "make_master.py", 'star', op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
            elif 'Max.xlsx' in file:
                op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
                subprocess.run(["python", "make_master.py", insname, op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
            elif 'inamdar.xlsx' in file:
                op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
                subprocess.run(["python", "make_master.py", insname, op, '', path])
                move_master_to_master_insurer('')
                print(f'processed {path}')
                flag = 1
                break
        if flag == 0:
            # code for 2nd condtion
            process_insurer_pdfs(folder_name, insname, files)
            pass
    pass


def process_insurer_pdfs(folder_name, insname, files):
    for f in files:
        if '.pdf' in f:
            fpath = directory + '/' + folder_name + '/' + f
            subprocess.run(["python", "make_insurer_excel.py", insname, fpath])
        pass
    pass


if __name__ == "__main__":
    process_values('01/08/2020', '18/09/2020', 'all')
    # a = check_and_download_attachment('12', 'Raksha', 'Max',
    #                               "Claim Settlement Letter From Raksha Health Insurance TPA Pvt.Ltd. (M58ADD676ILBS,9331198,Dr. Aditi Goel.)")
    pass
