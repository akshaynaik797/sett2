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
    if 'MEDICLAIM COMPUTATION' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'vidal' + hosp_name + '.xlsx'
    t, wq = 0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    for t in range(0, len(onlyfiles)):
        def select_sheet(wks, sheet_name):

            if sheet_name not in wks.sheetnames:
                wks.create_sheet(sheet_name)
            # print(sheet_name)

            wks.save('temp_files/vidal.xlsx')


        def select_column(wks, s, ro):
            sheet = wks.worksheets[s]

            max_col1 = sheet.max_column
            s = []
            for j in range(1, max_col1 + 1):
                cell_obj = sheet.cell(row=1, column=j)
                s.append(cell_obj.value)


        # print(s)

        # print(mypath+'attachments/'+onlyfiles[t])
        CCN = onlyfiles[t].replace('.pdf', '')
        # print(CCN)

        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all', line_scale=100)
        tables.export('temp_files/foo1.xlsx', f='excel')
        loc = "temp_files/foo1.xlsx"
        wb1 = openpyxl.load_workbook(loc)
        wb = xlrd.open_workbook(loc)

        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()

        sh1 = ['Sr No.', 'ccn', 'Claim No.', 'Claimant Name', 'Insured Person', 'Policy No.', 'Corporate Name',
               'Insurance Company', 'Emp no./Ref-no.', 'Diagnosis', 'TDS Amount', 'Co-pay Amt', 'Deductible Amt',
               'Discount allowed', 'Settled Amt', 'EFT No', 'Transaction date', 'DOA', 'DOD', 'IP No.', 'Auth.Amt',
               'Claim File', 'insurance company', 'Deductible Amt', 'net payable', 'settled amt']
        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]
        sh2 = ['Sr No.', 'ccn', 'Bill No.', 'Bill Date', 'Nature of  Expenditure', 'Amt Claimed',
               'Disallowed / Non Medical Expenses Rs.', 'Amount Settled Rs.', 'Remarks']
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]

        sheet_1 = wb.sheet_by_index(0)
        sheet_1.cell_value(0, 0)
        max_row = sheet_1.nrows
        max_col = sheet_1.ncols
        wd = wb1.worksheets[0]
        for i in range(2, max_col):
            if sheet_1.cell_value(1, i) == '':
                # print(i)
                wd.delete_cols(i + 1)
        wb1.save(loc)
        wb1.close()
        wb = xlrd.open_workbook(loc)
        sheet_1 = wb.sheet_by_index(0)
        sheet_1.cell_value(0, 0)
        max_row = sheet_1.nrows
        max_col = sheet_1.ncols
        hg = []
        b = []
        p = []
        np = []
        r = []
        rt = []
        mh = []
        m = 0
        for i in range(2, max_row):
            if sheet_1.cell_value(i, 1) == 'Total :':
                m = 1
                break
            hg.append(sheet_1.cell_value(i, 2))
            b.append(sheet_1.cell_value(i, 3))
            p.append(sheet_1.cell_value(i, 4))
            np.append(sheet_1.cell_value(i, 5))
            r.append(sheet_1.cell_value(i, 6))
            k = sheet_1.cell_value(i, 7)
            if k != '':
                x = k.find('.00')
                op1 = k[:x + 3]
                oy = k[x + 3:]
                rt.append(op1)
                mh.append(oy)
            else:
                rt.append(sheet_1.cell_value(i, 7))
                mh.append(sheet_1.cell_value(i, 8))

            # hg=[sub.replace('a.ii)', '') for sub in hg]

        if (m == 0):
            sheet_2 = wb.sheet_by_index(1)
            sheet_2.cell_value(0, 0)
            max_row = sheet_2.nrows
            for i in range(2, max_row):
                hg.append(sheet_2.cell_value(i, 2))
                b.append(sheet_2.cell_value(i, 3))
                p.append(sheet_2.cell_value(i, 4))
                np.append(sheet_2.cell_value(i, 5))
                r.append(sheet_2.cell_value(i, 6))
                rt.append(sheet_2.cell_value(i, 7))
                mh.append(sheet_1.cell_value(i, 8))
                if sheet_2.cell_value(i, 4) == 'Total':
                    break

        # print(hg)

        gh = []

        w5 = f.find('Claim No. :') + 11
        g = f[w5:]
        u3 = g.find('Claim File') + w5
        gh.append(f[w5:u3])

        w4 = f.find('Claimant Name :') + 15
        g = f[w4:]
        u2 = g.find('\n') + w4
        gh.append(f[w4:u2])

        w5 = f.find('Insured Name :') + 15
        g = f[w5:]
        u3 = g.find('Claimant Name :') + w5
        gh.append(f[w5:u3])

        w4 = f.find('Policy No. :') + 12
        g = f[w4:]
        u2 = g.find('Policy Start Date') + w4
        gh.append(f[w4:u2])

        w5 = f.find('Claim Settlement No.') + 21
        g = f[w5:]
        u3 = g.find('Settlement Date:') + w5
        ccn = f[w5:u3]

        w4 = f.find('Corporate Name') + 14
        g = f[w4:]
        u2 = g.find('Payee Name') + w4
        gh.append(f[w4:u2])

        w5 = f.find('Insurance Company :') + 20
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w4 = f.find('Emp no./Ref-no.') + 17
        g = f[w4:]
        u2 = g.find('IP No.') + w4
        gh.append(f[w4:u2])

        w5 = f.find('Diagnosis :') + 11
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w4 = f.find('TDS Amount') + 17
        g = f[w4:]
        u2 = g.find('\n') + w4
        gh.append(f[w4:u2])

        w5 = f.find('Co-pay Amt.') + 18
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Deductible Amt') + 20
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Discount allowed') + 24
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w4 = f.find('payment of Rs.') + 14
        g = f[w4:]
        u2 = g.find('vide') + w4
        gh.append(f[w4:u2])

        w5 = f.find('EFT No.') + 7
        g = f[w5:]
        u3 = g.find('dated') + w5
        gh.append(f[w5:u3])

        g = f[u3:]
        u2 = g.find('to the') + u3
        gh.append(f[u3 + 6:u2])
        gh = [sub.replace(':', '') for sub in gh]
        w5 = f.find('DOA :') + 5
        g = f[w5:]
        u3 = g.find('Hospital :') + w5
        gh.append(f[w5:u3])

        w5 = f.find('DOD :') + 5
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('IP No.') + 8
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Auth.Amt (Rs.)') + 17
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Claim File No. :') + 17
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Insurance Company :') + 20
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Deductible Amt (Rs.) :') + 23
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Total Approved (Rs.) :') + 23
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        w5 = f.find('Approval Amount (Rs) :') + 23
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])

        ccn = ccn.replace(':', '')
        gh = [sub.replace('  ', '') for sub in gh]
        # print(gh,ccn)

        for i in range(0, len(gh)):
            s1.cell(row=t + 2, column=1).value = t + 1
            s1.cell(row=t + 2, column=2).value = ccn
            s1.cell(row=t + 2, column=i + 3).value = gh[i]
        for i in range(0, len(hg)):
            row_num = s2.max_row + 1
            wq += 1
            s2.cell(row=row_num, column=1).value = wq
            s2.cell(row=row_num, column=2).value = ccn
            s2.cell(row=row_num, column=3).value = hg[i]
            s2.cell(row=row_num, column=4).value = b[i]
            s2.cell(row=row_num, column=5).value = p[i]
            s2.cell(row=row_num, column=6).value = np[i]
            s2.cell(row=row_num, column=7).value = r[i]
            s2.cell(row=row_num, column=8).value = rt[i]
            s2.cell(row=row_num, column=9).value = mh[i]
    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'vidal', op, '', wbkName])
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
