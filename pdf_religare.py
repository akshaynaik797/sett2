import json
import os
import subprocess
import sys

import openpyxl
import pdftotext
import tabula

from make_log import log_exceptions
from movemaster import move_master_to_master_insurer

try:
    pdfpath = sys.argv[1]
    onlyfiles = [os.path.split(pdfpath)[1]]
    mypath = os.path.dirname(pdfpath) + '/'

    hosp_name = ''

    with open(pdfpath, "rb") as f:
        pdf = pdftotext.PDF(f)
    with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
        f.write(" ".join(pdf))
    with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
        f = myfile.read()
    if 'Claim Approval & Settlement Letter' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'religare' + hosp_name + '.xlsx'
    t, wq = 0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('Sheet1')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]

    for t in range(0, len(onlyfiles)):
        sh1 = ['S.no', 'Claim No.', 'Policy No.', 'Bank Name', 'Proposer Name', 'Name of Patient',
               'Instrument/ NEFT No', 'Instrument/ NEFT Date', 'Date of admission', 'Date of Discharge', 'bill amt',
               'net amt', 'co-pay', 'deductible', 'Employee ID', 'Employee Name', 'hospital discount', 'Al no.',
               'total deduction']
        sh2 = ['S.no', 'Claim No.', 'Description', 'Rate per day/Quantity', 'No.of Day/Visits/Quantity', 'Bill Amount',
               'Admissible Amount', 'Deducted Amount', 'Deduction Reason']
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]
        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]

        print(onlyfiles[t])
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()

        hg = []
        w = f.find('Policy No.') + 10
        g = f[w:]
        u = g.find('\n') + w
        hg.append(f[w:u])

        w1 = f.find('Bank Name') + 10
        g = f[w1:]
        u1 = g.find('Successful') + w1
        hg.append(f[w1:u1])

        w2 = f.find('Proposer Name') + 14
        g = f[w2:]
        u2 = g.find('Policy No.') + w2
        u3 = g.find('\n') + w2
        u4 = g.find('Claimed') + w2
        hg.append(f[w2:u2] + f[u3:u4])

        w2 = f.find('Name of Patient') + 16
        g = f[w2:]
        u2 = g.find('\n') + w2
        u3 = g.find('Date of admission') + w2
        hg.append(f[w2:u2] + f[u2 + 60:u3])

        w3 = f.find('Instrument/ NEFT No') + 19
        g = f[w3:]
        u3 = g.find('\n') + w3
        hg.append(f[w3:u3])

        w4 = g.find('Instrument/ NEFT') + w3 + 17
        g = f[w4:]
        u4 = g.find('\n') + w4
        hg.append(f[w4:u4])

        w5 = f.find('Date of admission') + 18
        g = f[w5:]
        u5 = g.find('Date of Discharge') + w5
        hg.append(f[w5:u5])

        w6 = f.find('Date of Discharge') + 18
        g = f[w6:]
        u6 = g.find('\n') + w6
        hg.append(f[w6:u6])

        w6 = f.find('Bill Amount') + 18
        g = f[w6:]
        u6 = g.find('Instrument') + w6
        hg.append(f[w6:u6])

        w6 = f.find('Amount Paid') + 18
        g = f[w6:]
        u6 = g.find('Bank') + w6
        hg.append(f[w6:u6])

        if f.find('Co pay') != -1:
            w6 = f.find('Co pay') + 18
            g = f[w6:]
            u6 = g.find('Deductible') + w6
            hg.append(f[w6:u6])
        else:
            hg.append(' ')
        w6 = f.find('Deductible') + 18
        g = f[w6:]
        u6 = g.find('\n') + w6
        hg.append(f[w6:u6])

        w2 = f.find('Employee ID') + 14
        g = f[w2:]
        u3 = g.find('Employee Name') + w2
        hg.append(f[w2:u3])

        w2 = f.find('Employee Name') + 14
        g = f[w2:]
        u3 = g.find('Name of Proposer') + w2
        hg.append(f[w2:u3])

        w2 = f.find('Hospital Discount') + 25
        g = f[w2:]
        u3 = g.find('AL approved') + w2
        hg.append(f[w2:u3])

        w9 = f.find('AL No.') + 6
        g = f[w9:]
        u9 = g.find('\n') + w9 + 3
        ccn = (f[w9:u9])

        w9 = f.find('Claim No') + 15
        g = f[w9:]
        u9 = g.find('-') + w9 + 3
        hg.append(f[w9:u9])
        ccn = ccn.replace('\n', '')
        ccn = ccn.replace('.', '')
        ccn = ccn.replace(' ', '')
        if ccn == None:
            ccn = hg[-1]
        hg = [sub.replace('  ', '') for sub in hg]
        hg = [sub.replace(':', '') for sub in hg]
        hg = [sub.replace('\n', ' ') for sub in hg]
        s1.cell(row=t + 2, column=1).value = t + 1
        s1.cell(row=t + 2, column=2).value = ccn
        for i in range(0, len(hg)):
            s1.cell(row=t + 2, column=i + 3).value = hg[i]
        # df=tabula.read_pdf(mypath+onlyfiles[t],pages=2,multiple_tables=True,line_space=40)
        # tabula.convert_into(mypath + onlyfiles[t], 'temp_files/out.xls', output_format='excel', pages=2)
        tabula.convert_into(mypath + onlyfiles[t], 'temp_files/out' + str(t) + '.json', output_format='json',
                            pages='all')

        with open('temp_files/out' + str(t) + '.json') as f:
            data = json.load(f)
        k = 0
        l = 0
        u = 0
        w = 0
        mo = 0
        me = 0
        we = 0
        yt = 0
        p = []
        r = []
        ro = []
        po = []
        e = []
        eo = []
        kl = []
        lk = []
        for x in range(2, len(data)):
            d = data[x]["data"]
            # print(d)
            for i in d:
                m = [0, 0, 0, 0, 0, 0, 0, 0]
                for j in i:
                    for x, y in j.items():
                        h = str(y)
                        if (h.find('62.') != -1 and x == 'left'):
                            k = 1
                        if k == 1 and x == 'text':
                            k = 0
                            p.append(y)
                            m[0] = 1

                        if h.find('92.') != -1 and x == 'left':
                            u = 1
                        if u == 1 and x == 'text':
                            u = 0
                            r.append(y)
                            m[1] = 1

                        if h.find('152.') != -1 and x == 'left':
                            w = 1
                        if w == 1 and x == 'text':
                            w = 0
                            po.append(y)
                            m[2] = 1

                        if h.find('222.') != -1 and x == 'left':
                            l = 1
                        if l == 1 and x == 'text':
                            l = 0
                            ro.append(y)
                            m[3] = 1

                        if h.find('282.') != -1 and x == 'left':
                            mo = 1
                        if mo == 1 and x == 'text':
                            mo = 0
                            e.append(y)
                            m[4] = 1

                        if h.find('332.') != -1 and x == 'left':
                            me = 1
                        if me == 1 and x == 'text':
                            me = 0
                            eo.append(y)
                            m[5] = 1

                        if h.find('382.') != -1 and x == 'left':
                            we = 1
                        if we == 1 and x == 'text':
                            we = 0
                            kl.append(y)
                            m[6] = 1

                        if h.find('422.') != -1 and x == 'left':
                            yt = 1
                        if yt == 1 and x == 'text':
                            yt = 0
                            lk.append(y)
                            m[7] = 1

                if m[0] == 0:
                    p.append('')
                if m[1] == 0:
                    r.append(r[-1])
                if m[2] == 0:
                    po.append('')
                if m[3] == 0:
                    ro.append('')
                if m[4] == 0:
                    e.append('')
                if m[5] == 0:
                    eo.append('')
                if m[6] == 0:
                    kl.append('')
                if m[7] == 0:
                    lk.append('')
            # print(i)

        p = [sub.replace('\r', ' ') for sub in p]
        r = [sub.replace('\r', ' ') for sub in r]
        po = [sub.replace('\r', ' ') for sub in po]
        ro = [sub.replace('\r', ' ') for sub in ro]
        e = [sub.replace('\r', ' ') for sub in e]
        eo = [sub.replace('\r', ' ') for sub in eo]
        kl = [sub.replace('\r', ' ') for sub in kl]
        lk = [sub.replace('\r', ' ') for sub in lk]

        for i in range(0, len(p) - 1):
            if i == 0:
                continue
            if r[i] == 'Description':
                continue
            row_num = s2.max_row
            wq = wq + 1
            s2.cell(row=row_num + 1, column=1).value = wq
            s2.cell(row=row_num + 1, column=2).value = ccn
            s2.cell(row=row_num + 1, column=3).value = r[i]
            s2.cell(row=row_num + 1, column=4).value = ro[i]
            s2.cell(row=row_num + 1, column=5).value = po[i]
            s2.cell(row=row_num + 1, column=6).value = e[i]
            s2.cell(row=row_num + 1, column=7).value = eo[i]
            s2.cell(row=row_num + 1, column=8).value = kl[i]
            s2.cell(row=row_num + 1, column=9).value = lk[i]
        ded = kl[-1]
        s1.cell(row=t + 2, column=19).value = ded
    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'religare', op, '', wbkName])
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
