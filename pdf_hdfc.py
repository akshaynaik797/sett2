import subprocess
import openpyxl
import sys
import camelot
import pdftotext
import xlrd
from make_log import log_exceptions
from movemaster import move_master_to_master_insurer

try:
    pdfpath = sys.argv[1]
    hosp_name = ''

    with open(pdfpath, "rb") as f:
        pdf = pdftotext.PDF(f)
    with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
        f.write(" ".join(pdf))
    with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
        f = myfile.read()
    if 'Settlement Letter Without Prejudice' not in f:
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    t, wq = 0, 0
    wbkName = 'temp_files/' + 'hdfc' + hosp_name + '.xlsx'
    wbk = openpyxl.Workbook()
    wbk.create_sheet('1')
    wbk.create_sheet('2')
    s1 = wbk.worksheets[0]
    s2 = wbk.worksheets[1]
    s3 = wbk.worksheets[2]

    sh1 = ['Sr No.', 'CCN', 'HDFC ERGO ID', 'Patient Name', 'Policy No', 'Account No.', 'Bank name', 'Diagnosis',
           'Settled amount', 'Main Member', 'UTR No', 'Transaction Date', 'doa', 'dod']
    sh2 = ['Sr No.', 'CCN', 'Service Type', 'Claimed  Amount', 'Deduction  Amount', 'Discount', 'Settled  Amount',
           'Remarks']
    sh3 = ['Sr No.', 'CCN', 'Service Tax', 'Total with Service Tax', 'TDS', 'Co-Payment', 'Cheque Amount',
           'total discont', 'deductible', 'settled amount']
    for i in range(0, len(sh1)):
        s1.cell(row=1, column=i + 1).value = sh1[i]
    for i in range(0, len(sh2)):
        s2.cell(row=1, column=i + 1).value = sh2[i]
    for i in range(0, len(sh3)):
        s3.cell(row=1, column=i + 1).value = sh3[i]

    with open(pdfpath, "rb") as f:
        pdf = pdftotext.PDF(f)
    with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
        f.write(" ".join(pdf))
    with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
        f = myfile.read()

    gh = []

    w5 = f.find('HDFC ERGO ID : ') + 15
    g = f[w5:]
    u3 = g.find('\n') + w5
    gh.append(f[w5:u3])

    w4 = f.find('Patient Name :') + 14
    g = f[w4:]
    u2 = g.find('Main Member') + w4
    gh.append(f[w4:u2])

    w = f.find('policy number') + 13
    g = f[w:]
    u3 = g.find(',') + w
    gh.append(f[w:u3])

    w3 = f.find('Account No.') + 11
    g = f[w3:]
    u1 = g.find('with') + w3
    gh.append(f[w3:u1])

    u2 = g.find('and') + w3
    gh.append(f[u1 + 5:u2])

    w3 = f.find('Ailment :') + 9
    g = f[w3:]
    u1 = g.find('Hospitalization') + w3
    di = f[w3:u1]
    di = di.replace('HOSPITAL', '')
    gh.append(di)

    w = f.find('claim with CCN') + 14
    g = f[w:]
    u3 = g.find(',') + w
    ccn = f[w:u3]

    w3 = f.find('sum of') + 7
    g = f[w3:]
    u1 = g.find('(') + w3
    gh.append(f[w3:u1])

    w5 = f.find('Main Member') + 13
    g = f[w5:]
    u3 = g.find('\n') + w5
    gh.append(f[w5:u3])

    w5 = f.find('UTR') + 7
    g = f[w5:]
    u3 = g.find('and') + w5
    gh.append(f[w5:u3])

    w5 = f.find('Transaction Date') + 16
    g = f[w5:]
    u3 = g.find('\n') + w5
    gh.append(f[w5:u3])

    w5 = f.find('From :') + 7
    g = f[w5:]
    u3 = g.find('To') + w5
    gh.append(f[w5:u3])

    w5 = u3 + 4
    g = f[w5:]
    u3 = g.find('\n') + w5
    gh.append(f[w5:u3])
    gh = [sub.replace('Rs.', '') for sub in gh]
    gh = [sub.replace('  ', '') for sub in gh]
    gh = [sub.replace('.', '') for sub in gh]
    gh = [sub.replace('\n', '') for sub in gh]
    ccn = ccn.replace('  ', '')
    # print(gh,ccn)
    s1.cell(row=t + 2, column=1).value = t + 1
    s1.cell(row=t + 2, column=2).value = ccn

    for i in range(0, len(gh)):
        s1.cell(row=t + 2, column=i + 3).value = gh[i]

    tables = camelot.read_pdf(pdfpath, pages='all', line_scale=10)
    tables.export('temp_files/foo1.xls', f='excel')
    loc = ("temp_files/foo1.xls")
    wb = xlrd.open_workbook(loc)
    sheet_1 = wb.sheet_by_index(0)
    sheet_1.cell_value(0, 0)
    sheet_2 = wb.sheet_by_index(1)
    sheet_2.cell_value(0, 0)
    max_row = sheet_1.nrows
    hg = []
    b = []
    p = []
    np = []
    r = []
    rt = []
    m = 0
    for i in range(2, max_row):

        hg.append(sheet_1.cell_value(i, 2))
        b.append(sheet_1.cell_value(i, 3))
        p.append(sheet_1.cell_value(i, 4))
        np.append(sheet_1.cell_value(i, 5))
        r.append(sheet_1.cell_value(i, 6))
        rt.append(sheet_1.cell_value(i, 7))
        if sheet_1.cell_value(i, 2) == 'Total':
            m = 1
            sett = sheet_1.cell_value(i, 3)
            dis = sheet_1.cell_value(i, 5)
            ded = sheet_1.cell_value(i, 4)
            break
    hg = [sub.replace('a.ii)', '') for sub in hg]
    max_row = sheet_2.nrows
    max_col = sheet_2.ncols
    # print(max_col)
    if (m == 0 and max_col == 8):
        for i in range(1, max_row):
            hg.append(sheet_2.cell_value(i, 2))
            b.append(sheet_2.cell_value(i, 3))
            p.append(sheet_2.cell_value(i, 4))
            np.append(sheet_2.cell_value(i, 5))
            r.append(sheet_2.cell_value(i, 6))
            rt.append(sheet_2.cell_value(i, 7))
            if sheet_2.cell_value(i, 2) == 'Total':
                sett = sheet_2.cell_value(i, 3)
                dis = sheet_2.cell_value(i, 5)
                ded = sheet_2.cell_value(i, 4)
                # print(dis)
                break
    elif (m == 0 and max_col == 7):
        for i in range(1, max_row):
            x = sheet_2.cell_value(i, 1)
            if x[0] >= '0' and x[0] <= '9':
                w3 = x.find('\n') + 1
                x = x[w3:]
            # print(x)
            hg.append(x)
            b.append(sheet_2.cell_value(i, 2))
            p.append(sheet_2.cell_value(i, 3))
            np.append(sheet_2.cell_value(i, 4))
            r.append(sheet_2.cell_value(i, 5))
            rt.append(sheet_2.cell_value(i, 6))
            if sheet_2.cell_value(i, 1) == 'Total':
                sett = sheet_2.cell_value(i, 2)
                ded = sheet_2.cell_value(i, 3)
                dis = sheet_2.cell_value(i, 4)
                # print(dis)
                break
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
    # print(hg,b,p,np,r,rt)
    gh = []
    if f.find('Service Tax') != -1:
        w5 = f.find('Service Tax') + 12
        g = f[w5:]
        u3 = g.find('\n') + w5
        gh.append(f[w5:u3])
    else:
        gh.append(' ')
    if f.find('Total with Service Tax') != -1:
        w4 = f.find('Total with Service Tax') + 23
        g = f[w4:]
        u2 = g.find('\n') + w4
        gh.append(f[w4:u2])
    else:
        gh.append(' ')

    if f.find('TDS') != -1:
        w = f.find('TDS') + 4
        g = f[w:]
        u3 = g.find('\n') + w
        gh.append(f[w:u3])
    else:
        gh.append(' ')
    if f.find('Co-Payment') != -1:
        w = f.find('Co-Payment') + 10
        g = f[w:]
        u3 = g.find('\n') + w
        gh.append(f[w:u3])
    else:
        gh.append(' ')
    if f.find('Cheque Amount') != -1:
        w = f.find('Cheque Amount') + 13
        g = f[w:]
        u3 = g.find('\n') + w
        gh.append(f[w:u3])
    else:
        gh.append(' ')

    gh = [sub.replace('  ', '') for sub in gh]

    for i in range(0, len(gh)):
        s3.cell(row=t + 2, column=i + 3).value = gh[i]
    s3.cell(row=t + 2, column=1).value = t + 1
    s3.cell(row=t + 2, column=2).value = ccn
    s3.cell(row=t + 2, column=8).value = dis
    s3.cell(row=t + 2, column=9).value = ded
    s3.cell(row=t + 2, column=10).value = sett
    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'hdfc', op, '', wbkName])
    ###########################################################
    move_master_to_master_insurer('')
    print(f'processed {wbkName}')
    pass
except SystemExit as e:
    v = e.code
    if 'exit' in v:
        a =1
        os._exit(0)
except:
    log_exceptions()
    pass