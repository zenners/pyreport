from flask import Flask, request, jsonify, send_file
from flask_caching import Cache
import json
import requests
import pandas as pd
import numpy as np
import openpyxl
import flask_excel as excel
from io import BytesIO, StringIO
import os

import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime

cache = Cache(config={'CACHE_TYPE': 'simple'})

app = Flask(__name__)
excel.init_excel(app)
cache.init_app(app)
port = 5001
# port = int(os.getenv("PORT"))

def send_mail(send_from, send_to, subject, text, filename, server, port, username='', password='', isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filename, "rb").read())
    encoders.encode_base64(part)
    the_file = 'attachment; filename="{}"'.format(filename)
    part.add_header('Content-Disposition', the_file)
    msg.attach(part)

    # context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
    # SSL connection only working on Python 3+
    smtp = smtplib.SMTP(server, port)
    if isTls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


@app.route("/", methods=['GET'])
def index():
    return 'Hello World! I am running on port ' + str(port)

@app.route("/collectionreport", methods=['GET'])
def collectionreport():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    date = request.args.get('date')
    payload = {'date': date}

    url = 'https://rfc360-test.zennerslab.com/Service1.svc/collection'
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Mobile Number", "Loan Account No", "Customer Name", "Email", "FDD", "DD", "PNV", "MLV", "MI", "Term",
               "Sum of Penalty", "Amount Due", "Unpaid Months", "Paid Months", "OB", "Status", "Total Payment"]
    df = pd.DataFrame(data['collectionResult'])

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        totalPaymentsum = 0
        pnvsum = 0
        mlvsum = 0
        misum = 0
        sumOfPenaltysum = 0
        amountDuesum = 0
        outstandingBalancesum = 0
        df = pd.DataFrame(pd.np.empty((0, 18)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        # df['loanId'] = df['loanId'].astype(int)
        # df.sort_values(by=['loanId'], inplace=True)
        totalPaymentsum = pd.Series(df['totalPayment']).sum()
        pnvsum = pd.Series(df['pnv']).sum()
        mlvsum = pd.Series(df['mlv']).sum()
        misum = pd.Series(df['mi']).sum()
        sumOfPenaltysum = pd.Series(df['sumOfPenalty']).sum()
        amountDuesum = pd.Series(df['amountDue']).sum()
        outstandingBalancesum = pd.Series(df['outstandingBalance']).sum()
        # df['email'] = ''
        df = df[["loanId", "mobileNo", "loanAccountNo", "name", "email",  "fdd", "dd", "pnv", "mlv", "mi", "term",
                 "sumOfPenalty", "amountDue", "unapaidMonths", "paidMonths", "outstandingBalance", "status",
                 "totalPayment"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Collections", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})

    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Collections"]
    worksheet.merge_range('A1:R1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:R2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:R3', 'Collection Report', merge_format3)
    worksheet.merge_range('A4:R4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:R{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('F{}:G{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('H{}'.format(count + 1), pnvsum, merge_format4)
    worksheet.write('I{}'.format(count + 1), mlvsum, merge_format4)
    worksheet.write('J{}'.format(count + 1), misum, merge_format4)
    worksheet.write('L{}'.format(count + 1), sumOfPenaltysum, merge_format4)
    worksheet.write('M{}'.format(count + 1), amountDuesum, merge_format4)
    worksheet.write('P{}'.format(count + 1), outstandingBalancesum, merge_format4)
    worksheet.write('R{}'.format(count + 1), totalPaymentsum, merge_format4)
    worksheet.merge_range('A{}:R{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:R{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:R{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)
    writer.close()
    output.seek(0)

    print('sending spreadsheet')

    filename = "Collection Report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/accountingAgingReport", methods=['GET'])
def accountingAgingReport():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    date = request.args.get('date')
    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport"
    r = requests.post(url, json=payload)
    data = r.json()
    # sortData = sorted(data, key=lambda d: d['fullName'], reverse=False)

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Collector", "Customer Name", "Mobile Number", "Address", "Loan Account Number", "Today", "1-30",
               "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "Total", "Matured",
               "Due Principal", "Due Interest", "Due Penalty"]
    df = pd.DataFrame(data)

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        totalsum = 0
        principalsum = 0
        interestsum = 0
        penaltysum = 0
        df = pd.DataFrame(pd.np.empty((0, 19)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        totalsum = pd.Series(df['total']).sum()
        principalsum = pd.Series(df['principal']).sum()
        interestsum = pd.Series(df['interest']).sum()
        penaltysum = pd.Series(df['penalty']).sum()
        df = df[["collector", "fullName", "mobile", "address", "loanAccountNumber", "today","1-30", "31-60", "61-90",
                 "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "matured", "principal",
                 "interest", "penalty"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})

    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:S1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:S2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:S3', 'Aging Report (Accounting)', merge_format3)
    worksheet.merge_range('A4:S4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:S{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('M{}:N{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('O{}'.format(count + 1), totalsum, merge_format4)
    worksheet.write('Q{}'.format(count + 1), principalsum, merge_format4)
    worksheet.write('R{}'.format(count + 1), interestsum, merge_format4)
    worksheet.write('S{}'.format(count + 1), penaltysum, merge_format4)
    worksheet.merge_range('A{}:S{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:S{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:S{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Accounting) as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/operationAgingReport", methods=['GET'])
@cache.cached(timeout=50, key_prefix='all_comments')
def operationAgingReport():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    date = request.args.get('date')
    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Mobile Number", "Address", "Term", "FDD", "Status",
               "PNV", "MLV", "bPNV", "bMLV", "MI", "Not Due", "Matured", "Today", "1-30", "31-60", "61-90", "91-120",
               "121-150", "151-180", "181-360", "360 & over", "Total", "Due Principal", "Due Interest", "Due Penalty"]
    df = pd.DataFrame(data['operationAgingReportJson'])
    # df['appId'] = df['appId'].astype(int)
    # df.sort_values(by=['appId'])

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        totalsum = 0
        principalsum = 0
        interestsum = 0
        penaltysum = 0
        df = pd.DataFrame(pd.np.empty((0, 28)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        totalsum = pd.Series(df['total']).sum()
        principalsum = pd.Series(df['duePrincipal']).sum()
        interestsum = pd.Series(df['dueInterest']).sum()
        penaltysum = pd.Series(df['duePenalty']).sum()
        df = df[["appId", "loanaccountNumber", "fullName", "mobile", "address", "term", "fdd", "status", "PNV",
                 "MLV", "bPNV", "bMLV", "mi", "notDue", "matured", "today", "1-30", "31-60", "61-90", "91-120",
                 "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty"]]
        print(df)
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:AB1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:AB2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:AB3', 'Aging Report (Operations)', merge_format3)
    worksheet.merge_range('A4:AB4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:AB{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('W{}:X{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('Y{}'.format(count + 1), totalsum, merge_format4)
    worksheet.write('Z{}'.format(count + 1), principalsum, merge_format4)
    worksheet.write('AA{}'.format(count + 1), interestsum, merge_format4)
    worksheet.write('AB{}'.format(count + 1), penaltysum, merge_format4)
    worksheet.merge_range('A{}:AB{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:AB{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:AB{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Operations) as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmemoreport2", methods=['GET'])
def newmemoreport2():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport"
    # url = "http://localhost:6999/reports/memoreport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Sub Product", "Memo Type", "Purpose", "Amount",
               "Status", "Date", "Created By", "Approved By", "Approved Remarks"]

    creditDf = pd.DataFrame(data['Credit'])
    if creditDf.empty:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = 'Nothing to display'
        sumCredit = 0
        creditDf = pd.DataFrame(pd.np.empty((0, 12)))
        # return jsonify(greater_than_zero)
    else:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = ''
        sumCredit = pd.Series(creditDf['amount']).sum()
        creditDf.sort_values(by=['appId'], inplace=True)
        creditDf = creditDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                         "status", "date", "createdBy", "approvedBy", "approvedRemark"]]

    debitDf = pd.DataFrame(data['Debit'])
    if debitDf.empty:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = 'Nothing to display'
        sumDebit = 0
        debitDf = pd.DataFrame(pd.np.empty((0, 12)))
        # return jsonify(greater_than_zero)
    else:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = ''
        sumDebit = pd.Series(debitDf['amount']).sum()
        debitDf.sort_values(by=['appId'], inplace=True)
        debitDf = debitDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                           "status", "date", "createdBy", "approvedBy", "approvedRemark"]]


    creditDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Credit", header=headers)
    debitDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Debit", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
    # xldate_header = "Today"

    worksheetCredit = writer.sheets["Credit"]
    worksheetCredit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetCredit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheetCredit.merge_range('A3:L3', 'Memo Report(Credit)', merge_format3)
    worksheetCredit.merge_range('A4:L4', xldate_header, merge_format1)
    worksheetCredit.merge_range('A{}:L{}'.format(countCredit - 1, countCredit - 1), nodisplayCredit, merge_format1)
    worksheetCredit.merge_range('E{}:F{}'.format(countCredit + 1, countCredit + 1), 'TOTAL AMOUNT', merge_format3)
    worksheetCredit.write('G{}'.format(countCredit + 1), sumCredit, merge_format4)
    worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 3, countCredit + 3), 'Report Generated By :', merge_format2)
    worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 4, countCredit + 5), name, merge_format2)
    worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 7, countCredit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    worksheetDebit = writer.sheets["Debit"]
    worksheetDebit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetDebit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheetDebit.merge_range('A3:L3', 'Memo Report(Debit)', merge_format3)
    worksheetDebit.merge_range('A4:L4', xldate_header, merge_format1)
    worksheetDebit.merge_range('A{}:L{}'.format(countDebit - 1, countDebit - 1), nodisplayDebit, merge_format1)
    worksheetDebit.merge_range('E{}:F{}'.format(countDebit + 1, countDebit + 1), 'TOTAL AMOUNT', merge_format3)
    worksheetDebit.write('G{}'.format(countDebit + 1), sumDebit, merge_format4)
    worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 3, countDebit + 3), 'Report Generated By :', merge_format2)
    worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 4, countDebit + 5), name, merge_format2)
    worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 7, countDebit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmemoreport", methods=['GET'])
def newmemoreport():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport"
    # url = "http://localhost:6999/reports/memoreport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Sub Product", "Memo Type", "Purpose", "Amount",
               "Status", "Date", "Created By", "Remarks", "Approved Date", "Approved By", "Approved Remarks"]

    creditDf = pd.DataFrame(data['Credit'])
    if creditDf.empty:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = 'No Data'
        sumCredit = 0
        creditDf = pd.DataFrame(pd.np.empty((0, 14)))
        # return jsonify(greater_than_zero)
    else:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = ''
        sumCredit = pd.Series(creditDf['amount']).sum()
        creditDf.sort_values(by=['appId'], inplace=True)
        creditDf = creditDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]

    debitDf = pd.DataFrame(data['Debit'])
    if debitDf.empty:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = 'No Data'
        sumDebit = 0
        debitDf = pd.DataFrame(pd.np.empty((0, 14)))
        # return jsonify(greater_than_zero)
    else:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = ''
        sumDebit = pd.Series(debitDf['amount']).sum()
        debitDf.sort_values(by=['appId'], inplace=True)
        debitDf = debitDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                           "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]


    creditDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Credit", header=headers)
    debitDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Debit", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
    # xldate_header = "Today"

    worksheetCredit = writer.sheets["Credit"]
    worksheetCredit.merge_range('A1:N1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetCredit.merge_range('A2:N2', 'RFC360 Kwikredit', merge_format1)
    worksheetCredit.merge_range('A3:N3', 'Memo Report(Credit)', merge_format3)
    worksheetCredit.merge_range('A4:N4', xldate_header, merge_format1)
    worksheetCredit.merge_range('A{}:N{}'.format(countCredit - 1, countCredit - 1), nodisplayCredit, merge_format1)
    worksheetCredit.merge_range('E{}:F{}'.format(countCredit + 1, countCredit + 1), 'TOTAL AMOUNT', merge_format3)
    worksheetCredit.write('G{}'.format(countCredit + 1), sumCredit, merge_format4)
    worksheetCredit.merge_range('A{}:N{}'.format(countCredit + 3, countCredit + 3), 'Report Generated By :', merge_format2)
    worksheetCredit.merge_range('A{}:N{}'.format(countCredit + 4, countCredit + 5), name, merge_format2)
    worksheetCredit.merge_range('A{}:N{}'.format(countCredit + 7, countCredit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    worksheetDebit = writer.sheets["Debit"]
    worksheetDebit.merge_range('A1:N1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetDebit.merge_range('A2:N2', 'RFC360 Kwikredit', merge_format1)
    worksheetDebit.merge_range('A3:N3', 'Memo Report(Debit)', merge_format3)
    worksheetDebit.merge_range('A4:N4', xldate_header, merge_format1)
    worksheetDebit.merge_range('A{}:N{}'.format(countDebit - 1, countDebit - 1), nodisplayDebit, merge_format1)
    worksheetDebit.merge_range('E{}:F{}'.format(countDebit + 1, countDebit + 1), 'TOTAL AMOUNT', merge_format3)
    worksheetDebit.write('G{}'.format(countDebit + 1), sumDebit, merge_format4)
    worksheetDebit.merge_range('A{}:N{}'.format(countDebit + 3, countDebit + 3), 'Report Generated By :', merge_format2)
    worksheetDebit.merge_range('A{}:N{}'.format(countDebit + 4, countDebit + 5), name, merge_format2)
    worksheetDebit.merge_range('A{}:N{}'.format(countDebit + 7, countDebit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/memoreport", methods=['GET'])
def memoreport():
    output = BytesIO()

    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = 'https://rfc360-test.zennerslab.com/Service1.svc/getMemoReport'
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account No", "Full Name", "Mobile Number", "Sub Product", "Memo Type", "Purpose", "Amount",
               "Status", "Date Created", "Created By", "Remarks", "Approved Date", "Approved By", "Approved Remarks"]
    df = pd.DataFrame(data['getMemoReportResult'])
    df['loanId'] = df['loanId'].astype(int)
    df.sort_values(by=['loanId'], inplace=True)

    df = df[["loanId", "loanAccountNo", "fullName", "mobileNo", "subProduct", "memoType", "purpose", "amount",
             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
    # xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:O1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:O2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:O3', 'Memo Report', merge_format3)
    worksheet.merge_range('A4:O4', xldate_header, merge_format1)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/tat", methods=['GET'])
def tat():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/newtat"
    r = requests.post(url, json=payload)
    data = r.json()
    standard = data['standard']
    returned = data['return']

    standard_df = pd.read_csv(StringIO(standard))
    returned_df = pd.read_csv(StringIO(returned))

    if standard_df.empty:
        nodisplayStandard = 'No Data'
    else:
        nodisplayStandard = ''

    if returned_df.empty:
        nodisplayReturned = 'No Data'
    else:
        nodisplayReturned = ''

    countStandard = standard_df.shape[0] + 8
    countReturned = returned_df.shape[0] + 8


    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    standard_df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Standard")
    returned_df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Returned")

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
    # xldate_header = "Today"
    worksheetStandard = writer.sheets["Standard"]
    worksheetStandard.merge_range('A1:R1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetStandard.merge_range('A2:R2', 'RFC360 Kwikredit', merge_format1)
    worksheetStandard.merge_range('A3:R3', 'Turn Around Time Report (Standard)', merge_format3)
    worksheetStandard.merge_range('A4:R4', xldate_header, merge_format1)
    worksheetStandard.merge_range('A{}:R{}'.format(countStandard - 1, countStandard - 1), nodisplayStandard, merge_format1)
    # worksheetStandard.merge_range('E{}:F{}'.format(countStandard + 1, countStandard + 1), 'TOTAL AMOUNT', merge_format3)
    # worksheetStandard.write('G{}'.format(countStandard + 1), sumStandard, merge_format4)
    worksheetStandard.merge_range('A{}:R{}'.format(countStandard + 1, countStandard + 1), 'Report Generated By :',
                                  merge_format2)
    worksheetStandard.merge_range('A{}:R{}'.format(countStandard + 2, countStandard + 3), name, merge_format2)
    worksheetStandard.merge_range('A{}:R{}'.format(countStandard + 5, countStandard + 5),
                                  'Date & Time Report Generation ({})'.format(dateNow),
                                  merge_format2)
    worksheetReturned = writer.sheets["Returned"]
    worksheetReturned.merge_range('A1:W1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetReturned.merge_range('A2:W2', 'RFC360 Kwikredit', merge_format1)
    worksheetReturned.merge_range('A3:W3', 'Turn Around Time Report (Returned)', merge_format3)
    worksheetReturned.merge_range('A4:W4', xldate_header, merge_format1)
    worksheetReturned.merge_range('A{}:W{}'.format(countReturned - 1, countReturned - 1), nodisplayReturned, merge_format1)
    # worksheetReturned.merge_range('E{}:F{}'.format(countReturned + 1, countReturned + 1), 'TOTAL AMOUNT', merge_format3)
    # worksheetReturned.write('G{}'.format(countReturned + 1), sumReturned, merge_format4)
    worksheetReturned.merge_range('A{}:W{}'.format(countReturned + 1, countReturned + 1), 'Report Generated By :',
                                  merge_format2)
    worksheetReturned.merge_range('A{}:W{}'.format(countReturned + 2, countReturned + 3), name, merge_format2)
    worksheetReturned.merge_range('A{}:W{}'.format(countReturned + 5, countReturned + 5),
                                  'Date & Time Report Generation ({})'.format(dateNow),
                                  merge_format2)

    writer.close()
    output.seek(0)

    filename = "TAT {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/oldtat", methods=['GET'])
def oldtat():
    output = BytesIO()

    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/newtat"
    r = requests.post(url, json=payload)
    data = r.json()
    standard = data['standard']
    returned = data['return']

    standard_df = pd.read_csv(StringIO(standard))
    returned_df = pd.read_csv(StringIO(returned))

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    standard_df.to_excel(writer, sheet_name="Standard", index=False)
    returned_df.to_excel(writer, sheet_name="Returned", index=False)

    writer.close()
    output.seek(0)

    filename = "TAT {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/unappliedbalances", methods=['GET'])
def get_uabalances():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    url = "https://rfc360-test.zennerslab.com/Service1.svc/accountDueReportJSON"
    r = requests.post(url)
    data = r.json()

    # print(data)
    greater_than_zero = list(filter(lambda x: x['unappliedBalance'] > 0, data['accountDueReportJSONResult']))

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Mobile Number", "Amount Due", "Due Date",
               "Unapplied Balance"]
    df = pd.DataFrame(data['accountDueReportJSONResult'])

    print('df result: ', df)

    if df.empty:
        sum = 0
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 7)))
    # return jsonify(greater_than_zero)
    else:
        nodisplay = ''
        count = df.shape[0] + 8
        df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df['loanId'] = df['loanId'].astype(int)
        df.sort_values(by=['loanId'], inplace=True)
        sum = pd.Series(df['unappliedBalance']).sum()
        df = df[["loanId", "loanAccountNo", "name", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:G1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:G2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:G3', 'Unapplied Balances Report', merge_format3)
    worksheet.merge_range('A4:G4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:G{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('D{}:F{}'.format(count + 1, count + 1), 'TOTAL UNAPPLIED TODAY', merge_format3)
    worksheet.write('G{}'.format(count + 1), sum, merge_format4)
    worksheet.merge_range('A{}:G{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)
    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Unapplied Balance.xlsx"
    return send_file(output, attachment_filename=filename, as_attachment=True)



@app.route("/dccr", methods=['GET'])
def get_data():
    output = BytesIO()
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjson"
    r = requests.post(url, json=payload)
    data_json = r.json()
    sortData = sorted(data_json['DCCRjsonResult'], key=lambda d: d['postedDate'], reverse=False)

    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Loan Account Number", "Customer Name", "Mobile Number", "OR Number", "OR Date", "Net Cash",
               "Payment Source"]
    df = pd.DataFrame(sortData)

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        sum = 0
        df = pd.DataFrame(pd.np.empty((0, 7)))
    # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df["customerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df['amount'] = df['amount'].astype(float)
        sum = pd.Series(df['amount']).sum()
        df = df[['loanAccountNo', 'customerName', 'mobileNo', 'orNo', "postedDate", "amount",
                 "paymentSource"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:G1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:G2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:G3', 'Daily Cash Collection Report', merge_format3)
    worksheet.merge_range('A4:G4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:G{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('D{}:E{}'.format(count + 1, count + 1), 'TOTAL AMOUNT CASH', merge_format3)
    worksheet.write('F{}'.format(count + 1), sum, merge_format4)
    worksheet.merge_range('A{}:G{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)
    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/newdccr", methods=['GET'])
def get_data1():
    output = BytesIO()
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjsonNew"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # data = json.load(json_data)

    sortData = sorted(data_json['DCCRjsonNewResult'], key=lambda d: d['postedDate'], reverse=False)
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df = pd.DataFrame(sortData)

    if df.empty:
        count = df.shape[0] + 9
        nodisplay = 'No Data'
        amountsum = 0
        cashsum = 0
        checksum = 0
        advancessum = 0
        principalsum = 0
        interestsum = 0
        penaltysum = 0
        df = pd.DataFrame(pd.np.empty((0, 13)))
    # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 9
        nodisplay = ''
        conditions = [(df['paymentSource'] == 'Check')]
        df['total'] = np.select(conditions, [df['paymentCheck']], default=df['amount'])
        diff = df['total'] - (df['paidPrincipal'] + df['paidInterest'] + df['paidPenalty'])
        df['advances'] = round(diff, 2)
        amountsum = pd.Series(df['total']).sum()
        cashsum = pd.Series(df['amount']).sum()
        checksum = pd.Series(df['paymentCheck']).sum()
        advancessum = pd.Series(df['advances']).sum()
        principalsum = pd.Series(df['paidPrincipal']).sum()
        interestsum = pd.Series(df['paidInterest']).sum()
        penaltysum = pd.Series(df['paidPenalty']).sum()
        df = df[['paymentSource', 'cci', 'orDate', 'orNo', 'checkDate', 'checkNo', 'loanAccountNo', 'customerName',
                 'total', 'amount', 'paymentCheck', 'paidPrincipal', 'paidInterest', 'paidPenalty', 'advances']]

    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    merge_format5 = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:O1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:O2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:O3', 'Daily Cash/Check Collection Report', merge_format3)
    worksheet.merge_range('A4:O4', xldate_header, merge_format1)

    worksheet.merge_range('A6:A7', 'Payment Type', merge_format5)
    worksheet.merge_range('B6:B7', 'CCI', merge_format5)
    worksheet.merge_range('C6:C7', 'OR Date', merge_format5)
    worksheet.merge_range('D6:D7', 'OR #', merge_format5)
    worksheet.merge_range('E6:E7', 'Check Date', merge_format5)
    worksheet.merge_range('F6:F7', 'Check #', merge_format5)
    worksheet.merge_range('G6:G7', 'Loan Account Number', merge_format5)
    worksheet.merge_range('H6:H7', 'Customer Name', merge_format5)
    worksheet.merge_range('I6:K6', 'AMOUNT', merge_format5)
    worksheet.write('I7', 'Total', merge_format5)
    worksheet.write('J7', 'Cash', merge_format5)
    worksheet.write('K7', 'Check', merge_format5)
    worksheet.merge_range('L6:O6', 'LOAN REPAYMENT', merge_format5)
    worksheet.write('L7', 'Principal', merge_format5)
    worksheet.write('M7', 'Interest', merge_format5)
    worksheet.write('N7', 'Penalty (5%)', merge_format5)
    worksheet.write('O7', 'Advances', merge_format5)

    worksheet.merge_range('A{}:O{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.write('H{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('I{}'.format(count + 1), amountsum, merge_format4)
    worksheet.write('J{}'.format(count + 1), cashsum, merge_format4)
    worksheet.write('K{}'.format(count + 1), checksum, merge_format4)
    worksheet.write('L{}'.format(count + 1), principalsum, merge_format4)
    worksheet.write('M{}'.format(count + 1), interestsum, merge_format4)
    worksheet.write('N{}'.format(count + 1), advancessum, merge_format4)
    worksheet.write('O{}'.format(count + 1), penaltysum, merge_format4)
    worksheet.merge_range('A{}:O{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:O{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:O{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)
    # the writer has done its job

    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/dccr2", methods=['GET'])
def get_data2():
    output = 'test.xlsx'
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)

    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjson"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # pandas to excel
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    headers = ["Loan Account Number", "Customer Name", "Mobile Number", "OR Number", "OR Date", "Net Cash",
               "Payment Source"]
    df = pd.DataFrame(data_json['DCCRjsonResult'])
    df = df[['loanAccountNo', 'customerName', 'mobileno', 'orNo', "postedDate", "amountApplied", "paymentSource"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:G1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:G2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:G3', 'Daily Cash Collection Report', merge_format3)
    worksheet.merge_range('A4:G4', xldate_header, merge_format1)

    # the writer has done its job
    writer.save()

    # go back to the beginning of the stream
    # output.seek(0)
    print('sending spreadsheet')
    send_mail("cu.michaels@gmail.com", "jantzen@thegentlemanproject.com", "hello", "helloworld", filename,
              'smtp.gmail.com', '587', 'cu.michaels@gmail.com', 'jantzen216')
    return 'ok'
    # return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/monthlyincome", methods=['GET'])
def get_monthly():
    output = BytesIO()
    date = request.args.get('date')
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    payload = {'date': date}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Penalty Paid",
               "Interest Paid", "Principal Paid", "Unapplied Balance", "Payment Amount", "OR Date", "OR Number"]
    df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])

    if df.empty:
        count = df.shape[0] + 8
        sumPenalty = 0
        sumInterest = 0
        sumPrincipal = 0
        sumUnapplied = 0
        total = 0
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 8)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['appId'] = df['appId'].astype(int)
        df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['appId'], inplace=True)
        sumPenalty = pd.Series(df['penaltyPaid']).sum()
        sumInterest = pd.Series(df['interestPaid']).sum()
        sumPrincipal = pd.Series(df['principalPaid']).sum()
        sumUnapplied = pd.Series(df['unappliedBalance']).sum()
        total = pd.Series(df['paymentAmount']).sum()
        df = df[['appId', 'loanAccountno', 'name', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
                 'paymentAmount', "orDate", "orNo"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:J1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:J2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:J3', 'Monthly Income Report', merge_format3)
    worksheet.merge_range('A4:J4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:J{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.write('C{}'.format(count + 1), 'TOTAL', merge_format3)
    worksheet.write('D{}'.format(count + 1), sumPenalty, merge_format4)
    worksheet.write('E{}'.format(count + 1), sumInterest, merge_format4)
    worksheet.write('F{}'.format(count + 1), sumPrincipal, merge_format4)
    worksheet.write('G{}'.format(count + 1), sumUnapplied, merge_format4)
    worksheet.write('H{}'.format(count + 1), total, merge_format4)
    worksheet.merge_range('A{}:J{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)


    # #the writer has done its job
    writer.close()

    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Monthly Income {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/booking", methods=['GET'])
def get_booking():
    output = BytesIO()

    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/bookingReportJs"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Sub Product", "PNV", "MLV", "Finance Fee",
               "GCLI", "Handling Fee", "Term", "Rate", "MI", "Booking Date", "Approval Date",
               "Application Date", "Branch"]
    df = pd.DataFrame(data_json['bookingReportJsResult'])

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        PNVsum = 0
        principalsum = 0
        interestsum = 0
        insurancesum = 0
        handlingFeesum = 0
        monthlyAmountsum = 0
        df = pd.DataFrame(pd.np.empty((0, 16)))
        # return jsonify(greater_than_zero)
    else:
        nodisplay = ''
        df['forreleasingdate'] = df.forreleasingdate.apply(lambda x: x.split(" ")[0])
        df['approvalDate'] = df.approvalDate.apply(lambda x: x.split(" ")[0])
        df['applicationDate'] = df.applicationDate.apply(lambda x: x.split(" ")[0])
        df["customerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df['loanId'] = df['loanId'].astype(int)
        df.sort_values(by=['loanId'], inplace=True)
        count = df.shape[0] + 8
        # df['monthlyAmount'] = df['monthlyAmount'].astype(float)
        PNVsum = pd.Series(df['PNV']).sum()
        principalsum = pd.Series(df['principal']).sum()
        interestsum = pd.Series(df['interest']).sum()
        insurancesum = pd.Series(df['insurance']).sum()
        handlingFeesum = pd.Series(df['handlingFee']).sum()
        monthlyAmountsum = pd.Series(df['monthlyAmount']).sum()
        df = df[['loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "principal", "interest", "insurance",
                 "handlingFee", "term", "actualRate", "monthlyAmount", "forreleasingdate", 'approvalDate',
                 'applicationDate', 'branch']]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})

    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:P1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:P2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:P3', 'Booking Report  ', merge_format3)
    worksheet.merge_range('A4:P4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:P{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.write('D{}'.format(count + 1), 'TOTAL', merge_format3)
    worksheet.write('E{}'.format(count + 1), PNVsum, merge_format4)
    worksheet.write('F{}'.format(count + 1), principalsum, merge_format4)
    worksheet.write('G{}'.format(count + 1), interestsum, merge_format4)
    worksheet.write('H{}'.format(count + 1), insurancesum, merge_format4)
    worksheet.write('I{}'.format(count + 1), handlingFeesum, merge_format4)
    worksheet.write('L{}'.format(count + 1), monthlyAmountsum, merge_format4)
    worksheet.merge_range('A{}:P{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:P{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:P{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)
    # #the writer has done its job
    writer.close()

    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Booking Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/incentive", methods=['GET'])
def get_incentive():
    output = BytesIO()
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/generateincentiveReportJSON"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Booking Date", "App ID", "Customer Name", "Referral Type", "SA", "Branch", "Loan Type",  "Term", "MLV", "PNV",
               "MI", "Referrer"]
    df = pd.DataFrame(data_json['generateincentiveReportJSONResult'])

    if df.empty:
        count = df.shape[0] + 8
        PNVsum = 0
        monthlyAmountsum = 0
        totalAmountsum = 0
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 12)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df["borrowerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df['loanId'] = df['loanId'].astype(int)
        df.sort_values(by=['agentName'], inplace=True)
        PNVsum = pd.Series(df['PNV']).sum()
        monthlyAmountsum = pd.Series(df['monthlyAmount']).sum()
        totalAmountsum = pd.Series(df['totalAmount']).sum()
        df = df[
            ['bookingDate', 'loanId', 'borrowerName', 'refferalType', "SA", "dealerName", "loanType", "term",
             "totalAmount", "PNV", "monthlyAmount", "agentName"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:L3', 'Sales Referral Report  ', merge_format3)
    worksheet.merge_range('A4:L4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:L{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('G{}:H{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('I{}'.format(count + 1), totalAmountsum, merge_format4)
    worksheet.write('J{}'.format(count + 1), PNVsum, merge_format4)
    worksheet.write('K{}'.format(count + 1), monthlyAmountsum, merge_format4)
    worksheet.merge_range('A{}:L{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:L{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:L{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()

    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Sales Referral Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/mature", methods=['GET'])
def get_mature():
    output = BytesIO()
    date = request.args.get('date')
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    payload = {'date': date}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/maturedLoanReport"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Mobile Number", "Term", "bMLV", "Last Due Date",
               "Last Payment", "No. of Unpaid Months", "Total Payment", "Total Past Due", "Outstanding Balance",
               "No. of Months from Maturity"]
    df = pd.DataFrame(data_json['maturedLoanReportResult'])

    if df.empty:
        count = df.shape[0] + 8
        bMLVsum = 0
        totalPaymentSum = 0
        monthlydueSum = 0
        outStandingBalanceSum = 0
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 13)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['monthlydue'] = df['monthlydue'].astype(float)
        df['outStandingBalance'] = df['outStandingBalance'].astype(float)
        df['loanId'] = df['loanId'].astype(int)
        df["fullName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        bMLVsum = pd.Series(df['bMLV']).sum()
        totalPaymentSum = pd.Series(df['totalPayment']).sum()
        monthlydueSum = pd.Series(df['monthlydue']).sum()
        outStandingBalanceSum = pd.Series(df['outStandingBalance']).sum()
        df = df[['loanId', 'loanAccountNo', 'fullName', "mobileno", "term", "bMLV", "lastDueDate", "lastPayment",
                 "unpaidMonths", "totalPayment", "monthlydue", "outStandingBalance", "matured"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)
    print(count)
    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:M1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:M2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:M3', 'Matured Loans Report  ', merge_format3)
    worksheet.merge_range('A4:M4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:M{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('D{}:E{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('F{}'.format(count + 1), bMLVsum, merge_format4)
    worksheet.write('J{}'.format(count + 1), totalPaymentSum, merge_format4)
    worksheet.write('K{}'.format(count + 1), monthlydueSum, merge_format4)
    worksheet.write('L{}'.format(count + 1), outStandingBalanceSum, merge_format4)
    worksheet.merge_range('A{}:M{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:M{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:M{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()

    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Matured Loans Report as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/duetoday", methods=['GET'])
def get_due():
    output = BytesIO()
    date = request.args.get('date')
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    payload = {'date': date}
    url = "https://rfc360-test.zennerslab.com/Service1.svc/dueTodayReport"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App ID", "Loan Account Number", "Customer Name", "Mobile Number", "Loan Type", "Due Today Term",
               "MI", "Total Past Due", "Unpaid Penalty", "Monthly Due", "Last Payment Date", "Last Payment Amount"]

    df = pd.DataFrame(data_json['dueTodayReportResult'])

    if df.empty:
        count = df.shape[0] + 8
        monthlyAmmortizationsum = 0
        monthduesum = 0
        unpaidPenaltysum = 0
        lastPaymentAmountsum = 0
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 12)))
        # return jsonify(greater_than_zero)
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['monthlyAmmortization'] = df['monthlyAmmortization'].astype(float)
        df['monthdue'] = df['monthdue'].astype(float)
        df['loanId'] = df['loanId'].astype(int)
        df["fullName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        monthlyAmmortizationsum = pd.Series(df['monthlyAmmortization']).sum()
        monthduesum = pd.Series(df['monthdue']).sum()
        unpaidPenaltysum = pd.Series(df['unpaidPenalty']).sum()
        lastPaymentAmountsum = pd.Series(df['lastPaymentAmount']).sum()
        df = df[
            ["loanId", "loanAccountNo", "fullName", "mobileno", "loanType", "term", "monthlyAmmortization",
             "monthdue", "unpaidPenalty", "monthlydue", "lastPayment", "lastPaymentAmount"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:L3', 'Due Today Report  ', merge_format3)
    worksheet.merge_range('A4:L4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:L{}'.format(count - 1, count - 1), nodisplay, merge_format1)
    worksheet.merge_range('E{}:F{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
    worksheet.write('G{}'.format(count + 1), monthlyAmmortizationsum, merge_format4)
    worksheet.write('H{}'.format(count + 1), monthduesum, merge_format4)
    worksheet.write('I{}'.format(count + 1), unpaidPenaltysum, merge_format4)
    worksheet.write('L{}'.format(count + 1), lastPaymentAmountsum, merge_format4)
    worksheet.merge_range('A{}:L{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:L{}'.format(count + 4, count + 5), name, merge_format2)
    worksheet.merge_range('A{}:L{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()

    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Due Today Report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=port)
