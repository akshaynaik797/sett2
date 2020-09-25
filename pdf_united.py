import os
import subprocess
import PyPDF2
import openpyxl
import sys
import camelot
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
    if 'wanted_text' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'united' + hosp_name + '.xlsx'
    t, wq =0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    wbk.create_sheet('2')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    s3 = wbk.worksheets[2]
    for t in range(0, len(onlyfiles)):
        s3.cell(row=1, column=1).value = 'Sr. No.'
        s3.cell(row=1, column=2).value = 'Claim No/URL No'
        s3.cell(row=1, column=3).value = 'Amount'
        s3.cell(row=1, column=4).value = 'remark'
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()
        # print(f)

        sh2 = ['Policy Number', 'Employee No.', 'Patient Name', 'Insurance Company', 'EFT Transfer Code', 'DOA', 'DOD',
               'Loc No', 'Diagnosis', 'LOC/AL Amount', 'Claim Amount', 'ChequeDate', 'primary benifiary']
        for i in range(0, len(sh2)):
            s1.cell(row=1, column=i + 3).value = sh2[i]
        sh1 = ['UHC Approved Hospital Amt', 'UHC Approved Employee Amt', 'Insurer Approved Hospital Amt',
               'Insurer Approved Employee Amt', 'Payable Amount Hospital Amt', 'Payable Amount Employee Amt', 'TDS']

        for i in range(0, len(sh1)):
            s2.cell(row=1, column=i + 3).value = sh1[i]
        hg = []
        w = f.find('Policy No.') + 10
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w1 = f.find('Employee No.') + 12
        g = f[w1:]
        u1 = g.find('\n') + w1
        hg.append(f[w1:u1])

        w2 = f.find('Patient') + 7
        g = f[w2:]
        u2 = g.find('\n') + w2
        hg.append(f[w2:u2])

        w3 = f.find('policyholder of the') + 19
        g = f[w3:]
        u3 = g.find(',') + w3
        hg.append(f[w3:u3])

        w4 = f.find('EFT Transfer Code') + 17
        g = f[w4:]
        u4 = g.find('\n') + w4
        hg.append(f[w4:u4])

        w5 = f.find('DOA') + 4
        g = f[w5:]
        u5 = g.find('-') + w5
        hg.append(f[w5:u5])

        w6 = f.find('DOD') + 4
        g = f[w6:]
        u6 = g.find('\n') + w6
        hg.append(f[w6:u6])

        w7 = f.find('Loc No') + 7
        g = f[w7:]
        u7 = g.find('\n') + w7
        hg.append(f[w7:u7])

        w8 = f.find('Disease') + 7
        g = f[w8:]
        u8 = g.find('\n') + w8
        hg.append(f[w8:u8])

        w8 = f.find('LOC/AL Amount') + 14
        g = f[w8:]
        u8 = g.find('\n') + w8
        hg.append(f[w8:u8])

        w8 = f.find('Claim Amount') + 13
        g = f[w8:]
        u8 = g.find('\n') + w8
        hg.append(f[w8:u8])

        w8 = f.find('/Date') + 6
        g = f[w8:]
        u8 = g.find('-') + w8
        hg.append(f[w8:u8])

        w2 = f.find('/Employee') + 10
        g = f[w2:]
        u2 = g.find('\n') + w2
        hg.append(f[w2:u2])

        w9 = f.find('Claim No/URL No') + 15
        g = f[w9:]
        u9 = g.find('Loc') + w9
        ccn = (f[w9:u9])
        r5 = ccn.find('/') + 1
        ccn = ccn[r5:]
        ccn = ccn.replace('  ', '')
        hg = [sub.replace('  ', '') for sub in hg]
        hg = [sub.replace(':', '') for sub in hg]
        hg = [sub.replace('\n', ' ') for sub in hg]
        for i in range(0, len(hg)):
            s1.cell(row=t + 2, column=i + 3).value = hg[i]
        gh = []
        w = f.find('UHC Approved')
        # print(w)
        g = f[w:]
        w1 = g.find(':') + w + 1
        u = g.find('.') + w + 3
        gh.append(f[w1:u])
        w1 = g.find('Employee Amt') + w + 14
        u1 = g.find('\n') + w
        gh.append(f[w1:u1])

        w2 = f.find('Insurer Approved')
        g = f[w2:]
        w3 = g.find(':') + w2 + 1
        u2 = g.find('.') + w2 + 3
        gh.append(f[w3:u2])
        # print(w3,u2)

        w3 = g.find('Employee Amt') + w2 + 14
        u3 = g.find('\n') + w2
        gh.append(f[w3:u3])

        w4 = f.find('Payable Amount')
        g = f[w4:]
        w5 = g.find(':') + w4 + 1
        u4 = g.find('.') + w4 + 3
        gh.append(f[w5:u4])

        w5 = g.find('Employee Amt') + w4 + 14
        u5 = g.find('\n') + w4
        gh.append(f[w5:u5])
        # print(gh)
        if f.find('TDS:') != -1:
            w5 = f.find('TDS:') + 4
            g = f[w5:]
            u5 = g.find(']') + w5
            gh.append(f[w5:u5])
        else:
            gh.append(' ')
        # print(gh)
        for i in range(0, len(gh)):
            gh[i] = gh[i].replace('', "")
            s2.cell(row=t + 2, column=i + 3).value = gh[i].replace('', "")
        for wd in wbk.worksheets[:2]:
            wd.cell(row=1, column=1).value = 'Sr. No.'
            wd.cell(row=1, column=2).value = 'Claim No/URL No'
            wd.cell(row=t + 2, column=1).value = t + 1
            wd.cell(row=t + 2, column=2).value = ccn

        w5 = f.find('Cheque No/Date')
        g = f[w5:]
        w6 = g.find('\n') + w5
        u5 = g.find('Encashment') + w5
        we = f[w6:u5]
        we = we.replace('Deductions/Remarks', '')
        we = we.replace('  ', '')
        we = we.replace(' +', '+')
        we = we.replace('\n', '$$')
        we = we.replace('$$$$', '$$')
        # print(we)
        op = we.split('$$+')
        op = [i.replace('$$', '') for i in op]
        # print(op)
        ro = []
        yo = []
        for j in op:
            i = j.find('/-')
            if i != -1:
                ro.append(j[:i])
                yo.append(j[i + 2:])
            else:
                ro.append('0')
                yo.append('-')

        for i in range(0, len(ro)):
            row_num = s3.max_row + 1
            if yo[i] == 'Advance()':
                continue
            else:
                s3.cell(row=row_num, column=2).value = ccn
                s3.cell(row=row_num, column=3).value = ro[i]
                s3.cell(row=row_num, column=4).value = yo[i]

    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'united', op, '', wbkName])
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