import os
import re
import subprocess
import sys

import camelot
import openpyxl
import pdftotext
import xlrd

from make_log import log_exceptions
from movemaster import move_master_to_master_insurer

try:
    pdfpath = sys.argv[1]
    onlyfiles = [os.path.split(pdfpath)[1]]
    mypath = os.path.dirname(pdfpath)+'/'

    hosp_name = ''

    with open(pdfpath, "rb") as f:
        pdf = pdftotext.PDF(f)
    with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
        f.write(" ".join(pdf))
    with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
        f = myfile.read()
    if 'claim Settlement for Claim' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'Universal_Sompo' + hosp_name + '.xlsx'
    t, wq =0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]

    for t in range(0, len(onlyfiles)):
        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all')
        tables.export('temp_files/foo1.xls', f='excel')
        loc = "temp_files/foo1.xls"
        wb = xlrd.open_workbook(loc)
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()
        sh1 = ['sr no', 'CCN', 'IP NO', 'Patient Name', 'doa', 'dod', 'diagnosis', 'Beneficiary Name', 'Acc No.',
               'Bank name', 'IFSC code', 'UTR No.', 'NEFT Date', 'BilledAmount', 'SettledAmount', 'TDS', 'NetPayable',
               'DiscountAmt', 'COPay', 'deduction', 'Cashless Authorized']
        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]
        sh2 = ['sr no', 'CCN', 'category', 'deduction', 'reason']
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]
        hg = []

        regex = r'(?<=Claim Registration Number:) *\d+'
        result = re.search(regex, f)
        if result:
            hg.append(result.group().strip())
        else:
            w = f.find('Claim No:') + 10
            g = f[w:]
            u = g.find('\n') + w
            hg.append(f[w:u])

        w = f.find('Patient IP NO:') + 14
        g = f[w:]
        u = g.find('Claimed Amount:') + w
        hg.append(f[w:u])

        w = f.find('Patient Name:') + 13
        g = f[w:]
        u = g.find('Approved Amount') + w
        hg.append(f[w:u])

        w = f.find('Date of Admission:') + 18
        g = f[w:]
        u = g.find('Co Pay Amount:') + w
        hg.append(f[w:u])

        w = f.find('Date of Discharge:') + 18
        g = f[w:]
        u = g.find('TDS Deducted:') + w
        hg.append(f[w:u])

        w = f.find('Ailment:') + 10
        g = f[w:]
        u = g.find('Amount not') + w
        hg.append(f[w:u])

        w = f.find('Beneficiary Name:') + 17
        g = f[w:]
        u = g.find('NEFT Date:') + w
        hg.append(f[w:u])

        w = f.find('Beneficiary Acc No:') + 19
        g = f[w:]
        u = g.find('UTR No:') + w
        hg.append(f[w:u])

        w = f.find('Bank Name:') + 10
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('IFSC Code:') + 10
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('UTR No:') + 7
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('NEFT Date:') + 10
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('Claimed Amount:') + 14
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('Approved Amount:') + 15
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('TDS Deducted:') + 12
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        regex = r'(?<=Paid Amount after) *\d+'
        result = re.search(regex, f)
        if result:
            hg.append(result.group().strip())
        else:
            w = f.find('Paid Amount after TDS') + 22
            g = f[w:]
            u = g.find('\n') + w
            hg.append(f[w:u])

        w = f.find('Discount Amount:') + 15
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('Co Pay Amount:') + 13
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('Amount not paid*:') + 16
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w = f.find('Cashless Authorized Amount') + 26
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        hg = [sub.replace('  ', '') for sub in hg]
        hg = [sub.replace(':', '') for sub in hg]

        # print(hg)

        for i in range(0, len(hg)):
            s1.cell(row=t + 2, column=1).value = t + 1
            s1.cell(row=t + 2, column=i + 2).value = hg[i]

        regex = r'(?<=Reason for Deduction)\r?\n[ \S\n]+(?=In case of any variance)'
        regex2 = r'(?P<category>[\S ]+[^\d](?=\d+.0{2}))(?P<deduction>\d+.0{2})(?P<reason>[ \S]+)'
        result = re.search(regex, f)
        if result:
            raw = result.group().strip()
            s2_data = [match.groupdict() for match in re.compile(regex2).finditer(raw)]

        for i in s2_data:
            row_num = s2.max_row
            s2.cell(row=row_num + 1, column=1).value = row_num
            s2.cell(row=row_num + 1, column=2).value = hg[0]
            s2.cell(row=row_num + 1, column=3).value = i['category'].strip()
            s2.cell(row=row_num + 1, column=4).value = i['deduction'].strip()
            s2.cell(row=row_num + 1, column=5).value = i['reason'].strip()


    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'Universal_Sompo', op, '', wbkName])
    ###########################################################
    move_master_to_master_insurer('')
    print(f'processed {wbkName}')
except SystemExit as e:
    v = e.code
    if 'exit' in v:
        a =1
        os._exit(0)
except:
    log_exceptions()
    pass