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
    wbkName = 'temp_files/' + 'reliance' + hosp_name + '.xlsx'
    t, wq =0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    for t in range(0, len(onlyfiles)):
        sh1 = ['Sr No.', 'Claim No', 'Claimant/Patient', 'policy number', 'Employee id', 'UHID no', 'Diagnosis', 'DOA',
               'DOD', 'Duration', 'transacion mode', 'amount', 'amt in words', 'trnsaction date', 'NEFT no.',
               'Total Admissible Amount', 'Total Co-payment', 'Total Discount', 'Total Service Tax',
               'TDS Amount (if applicable)', 'Net Pay Amount']
        sh2 = ['Sr No.', 'Claim ID', 'category', 'Billed Amt(Rs)', 'Approved Amt(Rs)', 'Deduction Amt(Rs)',
               'Reason of Deduction (If any)']

        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]
        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all', line_scale=20)
        tables.export('temp_files/foo1.xls', f='excel')
        loc = ("temp_files/foo1.xls")
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()

        wb = xlrd.open_workbook(loc)

        w = f.find('claim number') + 12
        g = f[w:]
        u = g.find('towards') + w
        ccn = f[w:u]

        hg = []
        sheet_1 = wb.sheet_by_index(0)
        sheet_2 = wb.sheet_by_index(1)
        sheet_1.cell_value(0, 0)
        sheet_2.cell_value(0, 0)
        hg.append(sheet_1.cell_value(2, 2))
        hg.append(sheet_1.cell_value(1, 4))
        hg.append(sheet_1.cell_value(1, 6))
        hg.append(sheet_1.cell_value(2, 4))
        hg.append(sheet_1.cell_value(2, 6))
        hg.append(sheet_1.cell_value(3, 2))
        hg.append(sheet_1.cell_value(3, 4))
        hg.append(sheet_1.cell_value(3, 6))
        hg.append(sheet_2.cell_value(2, 1))
        hg.append(sheet_2.cell_value(2, 2))
        hg.append(sheet_2.cell_value(2, 3))
        hg.append(sheet_2.cell_value(2, 4))
        hg.append(sheet_2.cell_value(2, 5))
        s1.cell(row=t + 2, column=1).value = t + 1
        s1.cell(row=t + 2, column=2).value = ccn
        for i in range(0, len(hg)):
            s1.cell(row=t + 2, column=i + 3).value = hg[i]

        jh = []
        gh = []
        h = []
        g = []
        hj = []
        for j in range(2, tables.n):
            sheet_n = wb.sheet_by_index(j)
            sheet_n.cell_value(0, 0)
            if sheet_n.cell_value(1, 1) != 'Total Admissible Amount':
                for i in range(2, sheet_n.nrows):
                    jh.append(sheet_n.cell_value(i, 1))
                    gh.append(sheet_n.cell_value(i, 2))
                    h.append(sheet_n.cell_value(i, 3))
                    g.append(sheet_n.cell_value(i, 4))
                    hj.append(sheet_n.cell_value(i, 5))
            else:
                break
        for i in range(0, len(jh)):
            wq += 1
            row_num = s2.max_row
            s2.cell(row=row_num + 1, column=1).value = wq
            s2.cell(row=row_num + 1, column=2).value = ccn
            s2.cell(row=row_num + 1, column=3).value = jh[i]
            s2.cell(row=row_num + 1, column=4).value = gh[i]
            s2.cell(row=row_num + 1, column=5).value = h[i]
            s2.cell(row=row_num + 1, column=6).value = g[i]
            s2.cell(row=row_num + 1, column=7).value = hj[i]
        hg = []
        sheet_n = wb.sheet_by_index(j)
        hg.append(sheet_n.cell_value(1, 2))
        hg.append(sheet_n.cell_value(2, 2))
        hg.append(sheet_n.cell_value(3, 2))
        hg.append(sheet_n.cell_value(4, 2))
        hg.append(sheet_n.cell_value(5, 2))
        hg.append(sheet_n.cell_value(6, 2))
        for i in range(0, len(hg)):
            s1.cell(row=t + 2, column=i + 16).value = hg[i]
    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'reliance', op, '', wbkName])
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