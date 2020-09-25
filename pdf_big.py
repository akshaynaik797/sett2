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
    if 'Intimation No' not in f:
        if 'Bene Code' in f:
            subprocess.run(["python", 'pdf_' + 'small' + ".py", sys.argv[1]])
            os._exit(1)
        sys.exit(f'{pdfpath} wrong pdf recieved, so not processed')
    else:
        if 'Balaji Medical' in f:
            op = 'Tpappg@maxhealthcare.com May@2020 outlook.office365.com Max PPT'
            hosp_name = 'Max'
        else:
            op = 'mediclaim@inamdarhospital.org Mediclaim@2019 imap.gmail.com inamdar hospital'
            hosp_name = 'inamdar'
    ###########################################################
    wbkName = 'temp_files/' + 'big' + hosp_name + '.xlsx'
    t, wq =0, 0
    wbk = openpyxl.Workbook()
    wbk.create_sheet('Sheet1')
    s1 = wbk.worksheets[0]
    s3 = wbk.worksheets[1]
    for t in range(0, len(onlyfiles)):
        for wd in wbk.worksheets:
            wd.cell(row=1, column=1).value = 'Sr. No.'
            wd.cell(row=1, column=2).value = 'INTIMATION NO'
        sh1 = ['Policy number', 'Diagnosis', 'DOA', 'DOD', 'Claimant Name', 'ICD Codes Desc', 'Total amount claimed',
               'Hospitalisation payable amount', 'Pre hospitalisation payable amount',
               'Post hospitalisation payable amount', 'Add on Benefit(Hospital Cash / Patient care)',
               'Total Claim Payable Amount', 'deducted']
        sh2 = ['Nature of Expenditure', 'Amount Claimed', '	Approve d Amount', 'Disallowance Reasons / Remarks']
        for i in range(0, len(sh1)):
            s1.cell(row=1, column=i + 3).value = sh1[i]
        for i in range(0, len(sh2)):
            s3.cell(row=1, column=i + 3).value = sh2[i]


        def select_sheet(wks, sheet_name):

            if not sheet_name in wks.sheetnames:
                wks.create_sheet(sheet_name)
            # print(sheet_name)

            wks.save('temp_files/star.xlsx')


        def select_column(wks, s, ro):
            sheet = wks.worksheets[s]

            max_col = sheet.max_column
            s = []
            for i in range(1, max_col + 1):
                cell_obj = sheet.cell(row=1, column=i)
                s.append(cell_obj.value)


        # print(s)

        # print(CCN)

        tables = camelot.read_pdf(mypath + onlyfiles[t], pages='all', Line_scale=100)
        # print(tables)
        with open(mypath + onlyfiles[t], "rb") as f:
            pdf = pdftotext.PDF(f)

        with open('temp_files/output.txt', 'w', encoding='utf-8') as f:
            f.write(" ".join(pdf))
        with open('temp_files/output.txt', 'r',  encoding='utf-8') as myfile:
            f = myfile.read()
        # print(data)

        gh = []
        sd = []
        w2 = f.find('Intimation No') + 21
        g = f[w2:]
        u2 = g.find('Bill') + w2 - 1
        c = f[w2:u2]
        cli = c.replace(' ', '')
        w = f.find('Policy No') + 9
        g = f[w:]
        u = u2 = g.find('\n') + w
        gh.append(f[w:u])
        # print(f[w:u])
        w1 = f.find('Diagnosis') + 19
        g = f[w1:]
        u1 = g.find(',') + w1
        gh.append(f[w1:u1])

        if gh[-1].find('\n') != -1:
            o = gh[-1]
            x = gh[-1].find('\n')

            gh[-1] = o[:x] + o[x + 33:]

        w = f.find('DOA') + 3
        g = f[w:]
        u = u2 = g.find('\n') + w
        gh.append(f[w:u])

        w1 = f.find('DOD') + 3
        g = f[w1:]
        u1 = g.find('\n') + w1
        gh.append(f[w1:u1])

        w = f.find('Claimant Name') + 13
        g = f[w:]
        u = u2 = g.find('Product Name') + w
        gh.append(f[w:u])

        w1 = f.find('ICD Codes Desc') + 14
        g = f[w1:]
        u1 = g.find(',') + w1
        gh.append(f[w1:u1])

        gh = [sub.replace('  ', '') for sub in gh]
        # print(gh)
        s1.cell(row=t + 2, column=1).value = t + 1
        s1.cell(row=t + 2, column=2).value = cli
        for i in range(0, len(gh)):
            s1.cell(row=t + 2, column=i + 3).value = gh[i]

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
        m = 0
        for i in range(3, max_row):
            if sheet_1.cell_value(i, 1) == 'Total':
                m = 1
                break
            hg.append(sheet_1.cell_value(i, 2))
            b.append(sheet_1.cell_value(i, 5))
            p.append(sheet_1.cell_value(i, 8))

            k = sheet_1.cell_value(i, 9)
            if (k[0:5] == 'Refer'):
                u = k[11:]
                # print(u)
                sheet_4 = wb.sheet_by_index(3)
                sheet_4.cell_value(0, 0)
                for rw in range(0, sheet_4.nrows):
                    ty = sheet_4.cell_value(rw, 1)
                    # print(ty)
                    if (ty == u):
                        np.append(sheet_4.cell_value(rw, 2))
            else:
                np.append(sheet_1.cell_value(i, 9))
        # Refer Note #1
        sheet_1.cell_value(0, 0)
        hg = [sub.replace('a.ii)', '') for sub in hg]
        max_row = sheet_2.nrows
        if (m == 0):
            for i in range(3, max_row):
                if sheet_2.cell_value(i, 1) == 'Total':
                    break
                hg.append(sheet_2.cell_value(i, 2))
                b.append(sheet_2.cell_value(i, 5))
                p.append(sheet_2.cell_value(i, 8))
                k = sheet_2.cell_value(i, 9)
                if (k[0:5] == 'Refer'):
                    u = k[11:]
                    # print(u)
                    sheet_4 = wb.sheet_by_index(3)
                    sheet_4.cell_value(0, 0)
                    for rw in range(0, sheet_4.nrows):
                        ty = sheet_4.cell_value(rw, 1)
                        # print(ty)
                        if (ty == u):
                            np.append(sheet_4.cell_value(rw, 2))
                else:
                    np.append(sheet_2.cell_value(i, 9))

        # print(hg)
        for i in range(0, len(hg)):
            row_num = s3.max_row + 1
            wq += 1
            s3.cell(row=row_num, column=1).value = wq
            s3.cell(row=row_num, column=2).value = cli
            s3.cell(row=row_num, column=3).value = hg[i]
            s3.cell(row=row_num, column=4).value = b[i]
            s3.cell(row=row_num, column=5).value = p[i]
            s3.cell(row=row_num, column=6).value = np[i]
        a = 0
        for wd in wb.sheets():
            if wd.cell_value(1, 1) == 'Section':
                break
            a += 1
        # print(a)
        sheet_2 = wb.sheet_by_index(a)
        r = []
        ro = []
        max_row = sheet_2.nrows
        for i in range(2, max_row):
            # r.append(sheet_2.cell_value(i,1))
            ro.append(sheet_2.cell_value(i, 2))
        for i in range(0, len(ro)):
            # s2.cell(row=row_num, column=3).value = r[i]
            s1.cell(row=t + 2, column=i + 9).value = ro[i]
        # print(ro)
        if (f.find('Total Deduction') != -1):
            w1 = f.find('Total Deduction') + 16
            g = f[w1:]
            u1 = g.find('\n') + w1
            ty = (f[w1:u1])
            # print(f[w1:u1])
            ty = ty.replace('  ', '')
            s1.cell(row=t + 2, column=15).value = ty
        if(f.find('Less: Hospital Discounts')!=-1):
            w1=f.find('Less: Hospital Discounts')+25
            g=f[w1:]
            u1=g.find('\n')+w1
            ty=(f[w1:u1])
            dis1=ty.replace('  ','')
        else:dis1=0
        if(f.find('Less: Network Hospital Discount')!=-1):
            w1=f.find('Less: Network Hospital Discount')+31
            g=f[w1:]
            u1=g.find('\n')+w1
            ty=(f[w1:u1])
            dis2=ty.replace('  ','')
        else:dis2=0
        s1.cell(row=t+2, column=16).value = float(dis1)+float(dis2)


    print("Done")
    wbk.save(wbkName)
    wbk.close()
    subprocess.run(["python", "make_master.py", 'star', op, '', wbkName])
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