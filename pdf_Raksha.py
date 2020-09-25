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
    wbkName = 'temp_files/' + 'Raksha' + hosp_name + '.xlsx'
    t, wq =0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    for t in range(0, len(onlyfiles)):
        sh1 = ['Sr No.', 'Claim No', 'Member Id', 'Dev/AgentCode', 'Insured/Employee', 'Claimant/Patient',
               'Claim Amount', 'Pass Amount', 'Net Payable Amount', 'Payment', 'Diagnosis',
               'Insurance Company Claim No.', 'Policy Number', 'Policy Period', 'Period of Hospitalization',
               'Corporate Name', 'Deduction', 'TDS*', 'Neft-Ref/Cheque No.', 'Neft-Ref/Cheque Date', 'Hospital Name']
        sh2 = ['Sr No.', 'Claim ID', 'category', 'Billed Amt(Rs)', 'Deduction Amt(Rs)', 'Approved Amt(Rs)',
               'Reason of Deduction (If any)']

        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 1).value = sh1[i]
        for i in range(0, len(sh2)):
            s2.cell(row=1, column=i + 1).value = sh2[i]
        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all')
        tables.export('temp_files/foo1.xls', f='excel')
        loc = ("temp_files/foo1.xls")
        wb = xlrd.open_workbook(loc)
        sheet_1 = wb.sheet_by_index(0)
        sheet_1.cell_value(0, 0)
        s = []
        for k in range(1, len(sh1)):
            c = 0
            for j in [1, 3]:
                for i in range(1, sheet_1.nrows):
                    if (sheet_1.cell_value(i, j) == sh1[k]):
                        # print(sh1[k],j)
                        if (sheet_1.cell_value(i, j) == 'Period of Hospitalization'):
                            if (j == 1):
                                temp = sheet_1.cell_value(i, j + 1)
                                temp1 = temp.find('DOD')
                                doa = temp[:temp1]
                                dod = temp[temp1 + 4:]
                            if (j == 3):
                                temp = sheet_1.cell_value(i, j + 1)
                                temp1 = temp.find('DOD')
                                doa = temp[4:temp1]
                                dod = temp[temp1 + 4:]
                        elif (j == 1):
                            c = 1
                            s.append(sheet_1.cell_value(i, 2))
                        elif (j == 3):
                            c = 1
                            s.append(sheet_1.cell_value(i, 4))
            if c == 0:
                s.append(' ')
        s.append(doa)
        s.append(dod)
        # print(s)
        s1.cell(row=t + 2, column=1).value = t + 1
        for i in range(0, len(s)):
            s1.cell(row=t + 2, column=i + 2).value = s[i]
        hg = []
        gh = []
        h = []
        g = []
        hj = []
        for j in range(1, tables.n):
            sheet_n = wb.sheet_by_index(j)
            sheet_n.cell_value(0, 0)
            for i in range(2, sheet_n.nrows):
                hg.append(sheet_n.cell_value(i, 1))
                gh.append(sheet_n.cell_value(i, 2))
                h.append(sheet_n.cell_value(i, 3))
                g.append(sheet_n.cell_value(i, 4))
                hj.append(sheet_n.cell_value(i, 5))
        for i in range(0, len(hg)):
            wq += 1
            row_num = s2.max_row
            s2.cell(row=row_num + 1, column=1).value = wq
            s2.cell(row=row_num + 1, column=2).value = s[0]
            s2.cell(row=row_num + 1, column=3).value = hg[i]
            s2.cell(row=row_num + 1, column=4).value = gh[i]
            s2.cell(row=row_num + 1, column=5).value = h[i]
            s2.cell(row=row_num + 1, column=6).value = g[i]
            s2.cell(row=row_num + 1, column=7).value = hj[i]


    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'Raksha', op, '', wbkName])
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