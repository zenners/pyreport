from flask import Flask, request, jsonify, send_file
import json
import requests
import pandas as pd
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

app = Flask(__name__)
excel.init_excel(app)
# port = 5001
port = int(os.getenv("PORT"))

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

    date = request.args.get('date')
    payload = {'date': date}

    url = 'https://api360.zennerslab.com/Service1.svc/collection'
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App Id", "Mobile No", "Loan Account No", "Full Name", "FDD", "DD", "PNV", "MLV", "MI", "TERM",
               "Sum of Penalty", "Amount Due", "Unpaid Months", "Paid Months", "OB", "Status", "Total Payment"]
    df = pd.DataFrame(data['collectionResult'])

    df = df[["loanId", "mobileNo", "loanAccountNo", "name",  "fdd", "dd", "pnv", "mlv", "mi", "term",
             "sumOfPenalty", "amountDue", "unapaidMonths", "paidMonths", "outstandingBalance", "status", "totalPayment"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Collections", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Collections"]
    worksheet.merge_range('A1:Q1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:Q2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:Q3', 'Collection Report', merge_format3)
    worksheet.merge_range('A4:Q4', xldate_header, merge_format1)

    writer.close()
    output.seek(0)

    print('sending spreadsheet')

    filename = "Collection Report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/accountingAgingReport", methods=['GET'])
def accountingAgingReport():
    output = BytesIO()

    date = request.args.get('date')
    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Collector", "Full Name", "Mobile Number", "Address", "Loan Account Number", "Today", "1-30", "31-60",
               "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "Total", "Matured", "Due Principal",
               "Due Interest", "Due Penalty"]
    df = pd.DataFrame(data)

    df = df[["collector", "fullName", "mobile", "address", "loanAccountNumber", "today", "1-30", "31-60", "61-90",
             "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "matured", "principal",
             "interest", "penalty"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:S1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:S2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:S3', 'Aging Report (Accounting)', merge_format3)
    worksheet.merge_range('A4:S4', xldate_header, merge_format1)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Accounting) {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/operationAgingReport", methods=['GET'])
def operationAgingReport():
    output = BytesIO()

    date = request.args.get('date')
    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App Id", "Loan Account Number", "Full Name", "Mobile Number", "Address", "Term", "FDD", "Status", "PNV",
               "MLV", "bPNV", "bMLV", "MI", "Not Due", "Matured", "Today", "1-30", "31-60", "61-90", "91-120",
               "121-150", "151-180", "181-360", "360 & over", "Total", "Due Principal", "Due Interest", "Due Penalty"]
    df = pd.DataFrame(data)

    df = df[["appId", "loanaccountNumber", "fullName", "mobile", "address", "term", "fdd", "status", "PNV",
             "MLV", "bPNV", "bMLV", "mi", "notDue", "matured", "today", "1-30", "31-60", "61-90", "91-120",
             "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "As of {}".format(date)
    # xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:AB1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:AB2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:AB3', 'Aging Report (Operations)', merge_format3)
    worksheet.merge_range('A4:AB4', xldate_header, merge_format1)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Operations) {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmemoreport", methods=['GET'])
def newmemoreport():
    output = BytesIO()

    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport"
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App Id", "Loan Account Number", "Full Name", "Sub Product", "Memo Type", "Purpose", "Amount", "Status",
               "Date", "Created By", "Approved By", "Approved Remarks"]

    creditDf = pd.DataFrame(data["Credit"])
    debitDf = pd.DataFrame(data["Debit"])

    creditDf = creditDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount", "status",
                         "date", "createdBy", "approvedBy", "approvedRemark"]]
    debitDf = debitDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount", "status",
                       "date", "createdBy", "approvedBy", "approvedRemark"]]

    creditDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Credit", header=headers)
    debitDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Debit", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
    # xldate_header = "Today"

    worksheetCredit = writer.sheets["Credit"]
    worksheetCredit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetCredit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheetCredit.merge_range('A3:L3', 'Memo Report(Credit)', merge_format3)
    worksheetCredit.merge_range('A4:L4', xldate_header, merge_format1)

    worksheetDebit = writer.sheets["Debit"]
    worksheetDebit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheetDebit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
    worksheetDebit.merge_range('A3:L3', 'Memo Report(Debit)', merge_format3)
    worksheetDebit.merge_range('A4:L4', xldate_header, merge_format1)

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

    url = 'https://api360.zennerslab.com/Service1.svc/getMemoReport'
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["App Id", "Loan Account No", "Full Name", "Mobile No", "Sub Product", "Memo Type", "Purpose", "Amount",
               "Status", "Date Created", "Created By", "Remarks", "Approved Date", "Approved By", "Approved Remarks"]
    df = pd.DataFrame(data['getMemoReportResult'])

    df = df[["loanId", "loanAccountNo", "fullName", "mobileNo", "subProduct", "memoType", "purpose", "amount",
             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]

    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
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
    url = "https://api360.zennerslab.com/Service1.svc/accountDueReportJSON"
    r = requests.post(url)
    data = r.json()

    # print(data)
    greater_than_zero = list(filter(lambda x: x['unappliedBalance'] > 0, data['accountDueReportJSONResult']))

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Loan Account Number", "Customer Name", "Mobile No.", "Amount Due", "Due Date", "Unapplied Balance"]
    df = pd.DataFrame(data['accountDueReportJSONResult'])
    print('df result: ', df)
    if df.empty:
        sum = 0
        count = df.shape[0] + 8
        nodisplay = 'Nothing to display'
        df = pd.DataFrame(pd.np.empty((0, 6)))
    # return jsonify(greater_than_zero)
    else:
        sum = pd.Series(df['unappliedBalance']).sum()
        count = df.shape[0] + 8
        df = df[["loanAccountNo", "name", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]
        nodisplay = ''
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "Today"

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:F1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:F2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:F3', 'Unapplied Balances Report', merge_format3)
    worksheet.merge_range('A4:F4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:F{}'.format(count - 1, count - 1), nodisplay , merge_format1)
    worksheet.merge_range('C{}:E{}'.format(count, count), 'TOTAL UNAPPLIED TODAY', merge_format3)
    worksheet.write('F{}'.format(count), sum, merge_format4)
    worksheet.merge_range('A{}:F{}'.format(count + 2, count + 2), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:F{}'.format(count + 3, count + 4), name, merge_format2)
    worksheet.merge_range('A{}:F{}'.format(count + 6, count + 6), 'Date & Time Report Generation ({})'.format(dateNow),
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
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    # dateStart = '2018-06-26 00:00'
    # dateEnd = '2018-06-26 23:59'
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://api360.zennerslab.com/Service1.svc/DCCRjson"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Loan Account Number", "First Name", "Customer Name", "Mobile No.", "OR Number", "OR Date", "Net Cash",
               "Payment Source"]
    df = pd.DataFrame(data_json['DCCRjsonResult'])
    df = df[['loanAccountNo', 'firstName', 'customerName', 'mobileNo', 'orNo', "postedDate", "amountApplied",
             "paymentSource"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:H1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:H2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:H3', 'Daily Cash Collection Report', merge_format3)
    worksheet.merge_range('A4:H4', xldate_header, merge_format1)

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
    url = "https://api360.zennerslab.com/Service1.svc/DCCRjson"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # pandas to excel
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    headers = ["Loan Account Number", "Customer Name", "Mobile No.", "OR Number", "OR Date", "Net Cash",
               "Payment Source"]
    df = pd.DataFrame(data_json['DCCRjsonResult'])
    df = df[['loanAccountNo', 'customerName', 'mobileno', 'orNo', "postedDate", "amountApplied", "paymentSource"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
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
    url = "https://api360.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
    r = requests.post(url, json=payload)
    data_json = r.json()

    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Application ID", "Loan Account Number", "Customer Name", "Penalty Paid",
               "Interest Paid", "Principal Paid", "Payment Amount"]
    df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])
    df['penaltyPaid'] = df['penaltyPaid'].astype(float)
    df['interestPaid'] = df['interestPaid'].astype(float)
    df['principalPaid'] = df['principalPaid'].astype(float)
    df['paymentAmount'] = df['paymentAmount'].astype(float)

    sum = pd.Series(df['paymentAmount']).sum()
    count = df.shape[0] + 8

    df = df[['appId', 'loanAccountno', 'name', "penaltyPaid", "interestPaid", "principalPaid", 'paymentAmount']]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    # cell_format = workbook.add_format({'bold': True, 'align': 'left'})
    # total_cell_format = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:G1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:G2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:G3', 'Monthly Income Report', merge_format3)
    worksheet.merge_range('A4:G4', xldate_header, merge_format1)
    worksheet.merge_range('D{}:F{}'.format(count, count), 'TOTAL MONTHLY INCOME', merge_format3)
    worksheet.write('G{}'.format(count), sum, merge_format4)
    worksheet.merge_range('A{}:G{}'.format(count + 2, count + 2), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 3, count + 4), name, merge_format2)
    worksheet.merge_range('A{}:G{}'.format(count + 6, count + 6), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)


    #
    # #the writer has done its job
    writer.close()
    #
    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Monthly Income {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/booking", methods=['GET'])
def get_booking():
    output = BytesIO()
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://api360.zennerslab.com/Service1.svc/bookingReportJs"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Application ID", "Loan Account Number", "Customer Name", "Subproduct", "PNV", "MLV", "Finance Fee",
               "GCLI", "Handling Fee", "Term", "Rate", "Monthly Amortization", "Booking Date", "Approval Date",
               "Application Date", "Branch"]
    df = pd.DataFrame(data_json['bookingReportJsResult'])
    df['forreleasingdate'] = df.forreleasingdate.apply(lambda x: x.split(" ")[0])
    df['approvalDate'] = df.approvalDate.apply(lambda x: x.split(" ")[0])
    df['applicationDate'] = df.applicationDate.apply(lambda x: x.split(" ")[0])
    df = df[['loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "principal", "interest", "insurance",
             "handlingFee", "term", "actualRate", "monthlyAmount", "forreleasingdate", 'approvalDate',
             'applicationDate', 'branch']]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    merge_format = workbook.add_format({'align': 'left'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:P1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format)
    worksheet.merge_range('A2:P2', 'RFC360 Kwikredit', merge_format)
    worksheet.merge_range('A3:P3', 'Booking Report  ', merge_format)
    worksheet.merge_range('A4:P4', xldate_header, merge_format)

    # #the writer has done its job
    writer.close()
    #
    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Booking report {}-{}.xlsx".format(dateStart, dateEnd)
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
    url = "https://api360.zennerslab.com/Service1.svc/generateincentiveReportJSON"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Application ID", "Full Name", "Referral Type", "SA", "Branch", "Term", "MLV", "PNV",
               "MI", "Promodiser Name"]
    df = pd.DataFrame(data_json['generateincentiveReportJSONResult'])

    count = df.shape[0] + 8

    df = df[
        ["loanId", "borrowerName", "refferalType", "SA", "dealerName", "term", "totalAmount", "PNV", "monthlyAmount",
         "agentName"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    # cell_format = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    # merge_format4 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:J1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:J2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:J3', 'Sales Referral Report  ', merge_format3)
    worksheet.merge_range('A4:J4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:J{}'.format(count, count), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 1, count + 2), name, merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 4, count + 4), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()
    #
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
    url = "https://api360.zennerslab.com/Service1.svc/maturedLoanReport"
    r = requests.post(url, json=payload)
    data_json = r.json()
    sortData = sorted(data_json['maturedLoanReportResult'], key=lambda d: d['loanId'], reverse=False)
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Application ID", "Loan Account Number", "Customer Name", "Mobile No.", "Term", "Last Due Date",
               "Last Payment", "No. of Unpaid Months", "Total Past Due", "Outstanding Balance"]
    df = pd.DataFrame(sortData)
    df['monthlydue'] = df['monthlydue'].astype(float)
    df['outStandingBalance'] = df['outStandingBalance'].astype(float)

    count = df.shape[0] + 8

    df = df[['loanId', 'loanAccountNo', 'fullName', "mobileno", "term", "lastDueDate", "lastPayment", "unpaidMonths",
             "monthlydue", "outStandingBalance"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    # cell_format = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    # merge_format4 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:J1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:J2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:J3', 'Matured Loans Report  ', merge_format3)
    worksheet.merge_range('A4:J4', xldate_header, merge_format1)
    worksheet.merge_range('A{}:J{}'.format(count, count), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 1, count + 2), name, merge_format2)
    worksheet.merge_range('A{}:J{}'.format(count + 4, count + 4), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()
    #
    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Matured Loans Report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/duetoday", methods=['GET'])
def get_due():
    output = BytesIO()
    date = request.args.get('date')
    name = request.args.get('name')
    now = datetime.datetime.now()
    dateNow = now.strftime("%Y-%m-%d %I:%M %p")
    payload = {'date': date}
    url = "https://api360.zennerslab.com/Service1.svc/dueTodayReport"
    r = requests.post(url, json=payload)
    data_json = r.json()
    # return r.text
    # pandas to excel
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["Application ID", "Loan Account Number", "Customer Name", "Mobile No.", "Loan Type", "Term",
               "MI", "Past Due", "Monthly Due"]

    df = pd.DataFrame(data_json['dueTodayReportResult'])
    df['monthlyAmmortization'] = df['monthlyAmmortization'].astype(float)
    df['monthdue'] = df['monthdue'].astype(float)

    sum = pd.Series(df['monthlyAmmortization']).sum()
    count = df.shape[0] + 8
    # last = df["monthlyAmmortization"].iloc[-1]
    # print(last)
    df = df[
        ["loanId", "loanAccountNo", "fullName", "mobileno", "loanType", "term", "monthlyAmmortization",
         "monthdue", "monthlydue"]]
    df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)

    workbook = writer.book
    # merge_format = workbook.add_format({'align': 'left'})
    # cell_format = workbook.add_format({'bold': True, 'align': 'left'})
    # total_cell_format = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'center'})
    merge_format1 = workbook.add_format({'align': 'center'})
    merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
    merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
    merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
    xldate_header = "As of {}".format(date)

    worksheet = writer.sheets["Sheet_1"]
    worksheet.merge_range('A1:I1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
    worksheet.merge_range('A2:I2', 'RFC360 Kwikredit', merge_format1)
    worksheet.merge_range('A3:I3', 'Due Today Report  ', merge_format3)
    worksheet.merge_range('A4:I4', xldate_header, merge_format1)
    worksheet.merge_range('D{}:F{}'.format(count, count), 'TOTAL DUE TODAY', merge_format3)
    worksheet.write('G{}'.format(count), sum, merge_format4)
    worksheet.merge_range('A{}:I{}'.format(count + 2, count + 2), 'Report Generated By :', merge_format2)
    worksheet.merge_range('A{}:I{}'.format(count + 3, count + 4), name, merge_format2)
    worksheet.merge_range('A{}:I{}'.format(count + 6, count + 6), 'Date & Time Report Generation ({})'.format(dateNow),
                          merge_format2)

    # #the writer has done its job
    writer.close()
    #
    # #go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Due Today report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=port)
