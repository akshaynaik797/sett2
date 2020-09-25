import imaplib
import sqlite3

from make_log import log_exceptions


def accept_data(date):
    try:
        q1 = f"select hos_id,emailsubject from updation_detail_log where date > '{date}' and fieldreadflag='X'"
        with sqlite3.connect("database1.db") as con:
            cur = con.cursor()
            cur.execute(q1)
            r = cur.fetchall()
            if r is not None:
                data = []
                for i in r:
                    record = read_from_delete(i[1], i[0]), i[0], i[1]
                    data.append(record)
                return data
            return None
    except:
        return None


def read_from_delete(subject, hospital):
    try:
        server, email_id, password, inbox = "", "", "", ""
        if 'Max' in hospital:
            server, email_id, password, inbox = "outlook.office365.com", "Tpappg@maxhealthcare.com", "Sept@2020", '"Deleted Items"'
        elif 'inamdar' in hospital:
            server, email_id, password, inbox = "imap.gmail.com", "mediclaim@inamdarhospital.org", "Mediclaim@2019", '"[Gmail]/Trash"'
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_id, password)
        mail.select(inbox)
        type, data = mail.search(None, f'(SUBJECT "{subject}")')
        mid = data[0]  # this is the list, get last element and assign it to mid
        if type == 'OK' and len(mid) != 0:
            return 'Yes'
        return 'No'
    except:
        log_exceptions()
        return 'No'


if __name__ == "__main__":
    subject = 'HDFC ERGO HEALTH - Cashless Claim Approved-RC-HS20-11197900'
    # read_from_delete(subject, 'Max')
    a = accept_data('16/09/2020 12:47:46')
    pass
