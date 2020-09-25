import os
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
    mypath = os.path.dirname(pdfpath) + '/'

    hosp_name = ''

    with open(pdfpath, "rb") as f:
        pdf = pdftotext.PDF(f)
    with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
        f.write(" ".join(pdf))
    with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
        f = myfile.read()
    if 'claim has been' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'vipul' + hosp_name + '.xlsx'
    t, wq = 0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    # wbk.create_sheet('Sheet3')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    for t in range(0, len(onlyfiles)):
        sh1 = ['Sr. No.', 'File No', 'Patient Name', 'Admin Date', 'Emp Code', 'Dis Date', 'Claim Amt',
               'Total Bill Amt', 'Deduction', 'Co-Pay Deduction', 'Excess Pay Deduction', 'Approved Amount', 'Discount',
               'Settled Amt', 'TDS Amount', 'Net Amount']
        sh2 = ['S.no', 'File No.', 'Deducted Amount', 'Deduction Reason']
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]
        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]
        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all', line_scale=90)
        tables.export('temp_files/foo1.xlsx', f='excel')
        loc = "temp_files/foo1.xlsx"
        wb = xlrd.open_workbook(loc)
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()

        gh = []
        # print(f)
        w4 = f.find('Patient') + 12
        u2 = f.find('Admin Date')
        gh.append(f[w4:u2])
        if gh[-1] == ' ':
            print(f)
        w5 = f.find('Admin Date') + 11
        u3 = f.find('Emp Code')
        gh.append(f[w5:u3])

        w5 = f.find('Emp Code') + 8
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w1 = f.find('Dis') + 9
        u1 = f.find('File No')
        gh.append(f[w1:u1])

        w1 = f.find('File No') + 8
        g = f[w1:]
        u1 = g.find('\n') + w1
        ccn = f[w1:u1]
        ccn = ccn.replace('  ', '')
        # print(gh)

        gh = [sub.replace('  ', '') for sub in gh]
        if gh[0] == ' ' or gh[0] == '':
            w4 = f.find('information.') + 12
            u2 = f.find('Patient')
            gh[0] = f[w4:u2]
        # print(gh[0])
        gh = [sub.replace('\n', '') for sub in gh]
        gh = [sub.replace('  ', '') for sub in gh]

        sheet_1 = wb.sheet_by_index(1)
        sheet_1.cell_value(0, 0)
        max_row = sheet_1.nrows
        # print(max_row)
        max_col = sheet_1.ncols
        hg = []
        b = []
        p = []
        np = []
        r = []
        rt = []
        mh = []
        ro = []
        bo = []
        po = []
        no = []
        for i in range(1, max_row):
            ro.append(sheet_1.cell_value(i, 1))
            hg.append(sheet_1.cell_value(i, 2))
            b.append(sheet_1.cell_value(i, 3))
            p.append(sheet_1.cell_value(i, 4))
            np.append(sheet_1.cell_value(i, 5))
            r.append(sheet_1.cell_value(i, 6))
            rt.append(sheet_1.cell_value(i, 7))
            mh.append(sheet_1.cell_value(i, 8))
            bo.append(sheet_1.cell_value(i, 9))
            po.append(sheet_1.cell_value(i, 10))
            no.append(sheet_1.cell_value(i, 11))
            pass

        pass
        hg = [sub.replace('  ', '') for sub in hg]
        b = [sub.replace('  ', '') for sub in b]
        p = [sub.replace('  ', '') for sub in p]
        np = [sub.replace('  ', '') for sub in np]
        r = [sub.replace('  ', '') for sub in r]
        rt = [sub.replace('  ', '') for sub in rt]
        mh = [sub.replace('  ', '') for sub in mh]
        ro = [sub.replace('  ', '') for sub in ro]
        bo = [sub.replace('  ', '') for sub in bo]
        po = [sub.replace('  ', '') for sub in po]
        no = [sub.replace('\n', ' ') for sub in no]
        no = [sub.replace('\t', ' ') for sub in no]
        # print(no)
        xt = no[-1]
        op1 = xt.split(';')
        # print(op)
        s1.cell(row=t + 2, column=1).value = t + 1
        s1.cell(row=t + 2, column=2).value = ccn
        ph = ro[-1] + ' ' + hg[-1] + ' ' + b[-1] + ' ' + p[-1] + ' ' + np[-1] + ' ' + r[-1] + ' ' + rt[-1] + ' ' + mh[
            -1] + ' ' + bo[-1] + ' ' + po[-1]
        temp = [int(s) for s in ph.split() if s.isdigit()]
        # print(temp)
        for i in range(0, len(gh)):
            s1.cell(row=t + 2, column=2).value = ccn
            s1.cell(row=t + 2, column=i + 3).value = gh[i]
        for i in range(0, len(temp)):
            s1.cell(row=t + 2, column=len(gh) + 3 + i).value = temp[i]
        for i in op1:
            row_num = s2.max_row + 1
            wq += 1
            s2.cell(row=row_num, column=1).value = wq
            s2.cell(row=row_num, column=2).value = ccn
            w5 = i.find('Rs.') + 3
            g = i[w5:]
            u3 = g.find(' ') + w5
            # print(i[w5:u3])
            s2.cell(row=row_num, column=3).value = (i[w5:u3])
            s2.cell(row=row_num, column=4).value = (i[u3:])
        tr = []
        w5 = f.find('UTR') + 14
        g = f[w5:]
        u3 = g.find('dated') + w5
        tr.append(f[w5:u3])
        u1 = g.find('.') + w5
        tr.append(f[u3 + 5:u1])
        tr = [sub.replace('NEFT-', '') for sub in tr]
        for i in range(0, len(tr)):
            s1.cell(row=t + 2, column=len(gh) + len(temp) + 3 + i).value = tr[i]

    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'vipul', op, '', wbkName])
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
