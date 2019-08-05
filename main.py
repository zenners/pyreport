from flask import Flask, request, jsonify, send_file
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
from datetime import datetime
from pytz import timezone
from dateutil.parser import parse
from array import *
import ast

from pdf import pdf_api

app = Flask(__name__)
app.register_blueprint(pdf_api, url_prefix='/pdf')

excel.init_excel(app)
# port = 5001
port = int(os.getenv("PORT"))

fmtDate = "%m/%d/%y"
fmtTime = "%I:%M %p"
now_utc = datetime.now(timezone('UTC'))
now_pacific = now_utc.astimezone(timezone('Asia/Manila'))

dateNow = now_pacific.strftime(fmtDate)
timeNow = now_pacific.strftime(fmtTime)

comNameStyle = {'font':'Gill Sans MT', 'font_size': '16', 'bold': True, 'align': 'left'}
docNameStyle = {'font':'Segeo UI', 'font_size': '8', 'bold': True, 'align': 'left'}
periodStyle = {'font':'Segeo UI', 'font_size': '8', 'align': 'left'}
ledgerDataStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'right'}
ledgerDataStyle2 = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'text_wrap': True}
ledgerNameStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'left'}
undStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'bold': True, 'underline': True}
generatedStyle = {'font':'Segeo UI', 'font_size': '8', 'align': 'right'}
headerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True}
textWrapHeader = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True}
entriesStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'bottom': 2, 'align': 'center'}
borderFormatStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'bottom': 2, 'align': 'center', 'num_format': '₱#,##0.00'}
ledgerNum = {'font':'Segeo UI', 'font_size': '7', 'bottom': 2, 'align': 'right', 'num_format': '#,##0.00'}
topBorderStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'top': 2, 'align': 'center'}
footerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'right', 'num_format': '₱#,##0.00'}
sumStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'right'}
centerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center'}
numFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'num_format': '#,##0.00'}
stringFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'left', 'num_format': '#,##0.00'}
defaultFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right'}
defaultUnderlineFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'bottom': 2}
ledgerHeader = {'font':'Segeo UI', 'font_size': '7', 'align': 'left', 'bottom': 2}
ledgerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'left', 'underline': True, 'valign': 'bottom'}

styles = {
    'font-family': 'Segoe UI',
    'font-size': '9px'
}

dfstyles = {
    'font-family': 'Segoe UI',
    'font-size': '9px',
    'text-align': 'right'
}

# LAMBDA-JS URL

lambdaUrl = "https://ia-lambda-test.mybluemix.net/{}" #lambda-pivotal-live
# lambdaUrl = "https://ia-lambda-test.cfapps.io/{}" #lambda-pivotal-test
# lambdaUrl = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/{}" #lambda-amazon-live
# lambdaUrl = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/{}" #lambda-amazon-test
# lambdaUrl = "http://localhost:6999/{}" #lambda-localhost

# URL
# bluemixUrl = "https://rfc360.mybluemix.net/{}" #rfc-bluemix-live
bluemixUrl = "https://rfc360-test.mybluemix.net/{}" #rfc-bluemix-test
serviceUrl = "https://rfc360-test.zennerslab.com/Service1.svc/{}" #rfc-service-test
# serviceUrl = "https://api360.zennerslab.com/Service1.svc/{}" #rfc-service-live
# serviceUrl = "http://localhost:3000/{}" #rfc-localhost


# url2 = "http://localhost:15021/Service1.svc/getCustomerLedger" #test-local

def workbookFormat(workbook, styleName):
    workbook_format = workbook.add_format(styleName)
    return workbook_format

def dataframeStyle(worksheet, range1, range2, count, count1, merge_format):
    for c in range(ord(range1), ord(range2) + 1):
        worksheet.set_column('{}{}:{}{}'.format(chr(c), count, chr(c), count1 - 1), None, merge_format)

def dfDateFormat(df, colDateName):
    df[colDateName] = pd.to_datetime(df[colDateName])
    df[colDateName] = df[colDateName].map(lambda x: x.strftime('%m/%d/%y') if pd.notnull(x) else '')
    return df[colDateName]

def astype(df, colName, type):
    df[colName] = df[colName].astype(type)
    return df[colName]

def startDateFormat(dateStart):
    dateStart_object = datetime.strptime(dateStart, '%m/%d/%Y')
    payloaddateStart = dateStart_object.strftime('%m/%d/%y')
    return payloaddateStart

def endDateFormat(dateEnd):
    dateStart_object = datetime.strptime(dateEnd, '%m/%d/%Y')
    payloaddateEnd= dateStart_object.strftime('%m/%d/%y')
    return payloaddateEnd

def alphabet(secondRange):
    alphaList = [chr(c) for c in range(ord('A'), ord(secondRange) + 1)]
    return alphaList

def numbers(numRange):
    number = [number + 1 for number in range(numRange)]
    return number

def columnWidth(list1, list2):
    list3 = [max(value) for value in zip(list1, list2)]
    return list3

def dfwriter(dfName, writer, count):
    dfName(writer, startrow=count, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

def workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName):

    merge_format1 = workbook.add_format(periodStyle)
    merge_format2 = workbook.add_format(docNameStyle)
    merge_format3 = workbook.add_format(comNameStyle)
    merge_format5 = workbook.add_format(generatedStyle)

    worksheet.merge_range('A1:{}1'.format(range1), '{}'.format(companyName), merge_format3)
    worksheet.merge_range('A2:{}2'.format(range1), '{}'.format(reportTitle), merge_format2)
    worksheet.merge_range('A3:{}3'.format(range1), xldate_header, merge_format1)
    worksheet.merge_range('A4:{}4'.format(range1), '{}'.format(branchName), merge_format1)

    worksheet.merge_range('{}1:{}1'.format(range2, range3), 'Date Generated: {}'.format(dateNow), merge_format5)
    worksheet.merge_range('{}2:{}2'.format(range2, range3), 'Generated By: {}'.format(name), merge_format5)
    worksheet.merge_range('{}3:{}3'.format(range2, range3), 'Time Generated: {}'.format(timeNow), merge_format5)
    worksheet.merge_range('{}4:{}4'.format(range2, range3), 'Page - of -', merge_format5)

def paymentTypeWorksheet(worksheet, counts, type, merge_format7):
    worksheet.merge_range('A{}:A{}'.format(counts + 4, counts + 5), '#', merge_format7)
    worksheet.merge_range('B{}:B{}'.format(counts + 4, counts + 5), 'OR DATE', merge_format7)
    worksheet.merge_range('C{}:C{}'.format(counts + 4, counts + 5), 'OR #', merge_format7)
    worksheet.merge_range('D{}:D{}'.format(counts + 4, counts + 5), '{}'.format(type), merge_format7)
    worksheet.merge_range('E{}:G{}'.format(counts + 4, counts + 4), 'AMOUNT', merge_format7)
    worksheet.write('E{}'.format(counts + 5), 'TOTAL', merge_format7)
    worksheet.write('F{}'.format(counts + 5), 'CASH', merge_format7)
    worksheet.write('G{}'.format(counts + 5), 'CHECK', merge_format7)

def totalPaymentType(worksheet, counts, nodisplay, merge_format2, merge_format6, merge_format8):

    worksheet.merge_range('E{}:G{}'.format(counts + 6, counts + 6), nodisplay, merge_format6)
    worksheet.merge_range('A{}:B{}'.format(counts + 7, counts + 7), 'TOTAL:', merge_format2)
    worksheet.merge_range('A{}:G{}'.format(counts + 8, counts + 8),'', merge_format8)

def sumPaymentType(worksheet, count1, count2, count3, merge_format4):
    for c in range(ord('E'), ord('G') + 1):
        worksheet.write('{}{}'.format(chr(c), count1 + 7), "=SUM({}{}:{}{})".format(chr(c), count2 + 6, chr(c), count3 + 5), merge_format4)

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
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("collection")
    print(url)
    r = requests.post(url, json=payload)
    data = r.json()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "AMT DUE", "FDD", "PNV", "MLV", "MI", "TERM", "PEN", "INT",
               "PRIN", "UNPAID MOS", "PAID MOS", "HF", "DST", "NOTARIAL", "GCLI", "OB", "STATUS", "TOTAL PAYMENT"]
    df = pd.DataFrame(data['collectionResult'])
    list1 = [len(i) for i in headers]
    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 21)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        astype(df, 'dd', int)
        astype(df, 'term', int)
        astype(df, 'unapaidMonths', int)
        astype(df, 'paidMonths', int)
        astype(df, 'loanId', int)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['hf'] = 0
        df['dst'] = 0
        df['notarial'] = 0
        df['gcli'] = 0
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'fdd')
        df = round(df, 2)
        df = df[["num","loanId", "loanAccountNo", "name", "amountDue", "fdd", "pnv", "mlv", "mi", "term",
                 "sumOfPenalty", "totalInterest", "totalPrincipal", "unapaidMonths", "paidMonths", "hf", "dst", "notarial", "gcli", "outstandingBalance", "status",
                 "totalPayment"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    # df = df.style.set_properties(**styles)
    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Collections", header=None)

    workbook = writer.book


    worksheet = writer.sheets["Collections"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'D', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'E', 'E', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'F', 'F', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'G', 'I', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'J', 'J', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'K', 'M', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'N', 'O', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'P', 'T', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'U', 'U', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'V', 'V', 8, count, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    worksheet.freeze_panes(7, 0)

    range1 = 'S'
    range2 = 'T'
    range3 = 'V'
    companyName = 'RFSC'
    reportTitle = 'COLLECTION SUMMARY'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))
    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        if(x == 'P'):
            worksheet.write('P7', 'HF', workbookFormat(workbook, headerStyle))
        elif(x == 'Q'):
            worksheet.write('Q7', 'DST', workbookFormat(workbook, headerStyle))
        elif(x == 'R'):
            worksheet.write('R7', 'NOTARIAL', workbookFormat(workbook, headerStyle))
        elif(x == 'S'):
            worksheet.write('S7', 'GCLI', workbookFormat(workbook, headerStyle))
        else:
            worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('P6:S6', 'UPFRONT CHARGES', workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:V{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    for c in range(ord('E'), ord('V') + 1):
        if (chr(c) == 'F'):
            worksheet.write('F{}'.format(count + 1), None)
        elif (chr(c) == 'J'):
            worksheet.write('J{}'.format(count + 1), None)
        elif (chr(c) == 'N'):
            worksheet.write('N{}'.format(count + 1), None)
        elif (chr(c) == 'O'):
            worksheet.write('N{}'.format(count + 1), None)
        elif (chr(c) == 'U'):
            worksheet.write('U{}'.format(count + 1), None)
        else:
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                            workbookFormat(workbook, footerStyle))

    writer.close()
    output.seek(0)

    print('sending spreadsheet')

    filename = "Collection Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/accountingAgingReport", methods=['GET'])
def accountingAgingReport():

    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')

    payload = {'date': date}

    url = lambdaUrl.format("reports/accountingAgingReport")
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME",
               "COLLECTOR", "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "BUCKET", "CURR. TODAY",
               "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "OVER 360"]

    agingp1headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME", "COLLECTOR",
                      "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "BUCKET", "CURR. TODAY"]
    agingp11headers = ["1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "OVER 360"]
    agingp2headers = ["PRINCPAL", "INTEREST", "PENALTY"]

    agingp1DF = pd.DataFrame(data)
    agingp2DF = pd.DataFrame(data).copy()

    agingp1list1 = [len(i) for i in headers]
    agingp2list1 = [len(i) for i in agingp2headers]

    if agingp1DF.empty:
        count1 = agingp1DF.shape[0] + 8
        agingp1nodisplay = 'No Data'
        agingp1DF = pd.DataFrame(pd.np.empty((0, 19)))
        agingp1list2 = agingp1list1
    else:
        count1 = agingp1DF.shape[0] + 8
        agingp1nodisplay = ''
        agingp1DF['num'] = numbers(agingp1DF.shape[0])
        astype(agingp1DF, 'term', int)
        astype(agingp1DF, 'expiredTerm', int)
        astype(agingp1DF, 'appId', int)
        astype(agingp1DF, 'runningPNV', float)
        astype(agingp1DF, 'runningMLV', float)
        astype(agingp1DF, 'monthlyInstallment', float)
        astype(agingp1DF, 'duePrincipal', float)
        astype(agingp1DF, 'dueInterest', float)
        astype(agingp1DF, 'duePenalty', float)
        astype(agingp1DF, 'notDue', float)
        astype(agingp1DF, 'monthDue', float)
        dfDateFormat(agingp1DF, 'fdd')
        dfDateFormat(agingp1DF, 'lastPaymentDate')
        agingp1DF['loanAccountNumber'] = agingp1DF['loanAccountNumber'].map(lambda x: x.lstrip("'"))
        agingp1DF['lastPaymentDate'] = agingp1DF.lastPaymentDate.apply(lambda x: x.split(" ")[0])
        agingp1DF['totalDue'] = agingp1DF['totalmiDue'] + agingp1DF['duePenalty']
        agingp1DF["newCustomerName"] = agingp1DF['lastName'] + ', ' + agingp1DF['firstName'] + ' ' + agingp1DF['middleName'] + ' ' + agingp1DF['suffix']
        # agingp1DF['totalDueBreakdon'] = agingp1DF['duePrincipal'] + agingp1DF['dueInterest'] + agingp1DF['duePenalty']
        agingp1DF['ob'] = agingp1DF['notDue'] + agingp1DF['monthDue']
        # agingp1DF['adv'] = '-'
        agingp1DF = round(agingp1DF, 2)
        agingp1DF = agingp1DF[["num", "channelName", "partnerCode", "outletCode", "appId", "loanAccountNumber", "newCustomerName",
                               "alias", "fdd", "lastPaymentDate", "term", "expiredTerm", "monthlyInstallment", "stats", "ob", "runningMLV", "bucketing", "today",
                               "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty", "amountSum"]]
        agingp1list2 = [max([len(str(s)) for s in agingp1DF[col].values]) for col in agingp1DF.columns]

    # agingp1DF = agingp1DF.style.set_properties(**styles)
    agingp1DF.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="AgingP1", header=None)

    workbook = writer.book

    worksheetAgingP1 = writer.sheets["AgingP1"]

    dataframeStyle(worksheetAgingP1, 'A', 'A', 8, count1, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetAgingP1, 'B', 'D', 8, count1, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetAgingP1, 'E', 'E', 8, count1, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetAgingP1, 'F', 'H', 8, count1, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetAgingP1, 'I', 'J', 8, count1, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetAgingP1, 'K', 'L', 8, count1, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetAgingP1, 'M', 'M', 8, count1, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetAgingP1, 'N', 'N', 8, count1, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetAgingP1, 'O', 'P', 8, count1, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetAgingP1, 'Q', 'Q', 8, count1, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetAgingP1, 'R', 'Z', 8, count1, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AA8:AA{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AB8:AB{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AC8:AC{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AD8:AD{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AE8:AE{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(agingp1list1, agingp1list2)):
        worksheetAgingP1.set_column(col_num, col_num, value)

    worksheetAgingP1.freeze_panes(7, 0)

    def alphabetRange(firstRange, secondRange):
        alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
        return alphaList

    range1 = 'AA'
    range2 = 'AB'
    range3 = 'AE'
    companyName = 'RFSC'
    reportTitle = 'AGING REPORT'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheetAgingP1, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in agingp1headers]
    headersList1 = [i for i in agingp11headers]

    for x, y in zip(alphabetRange('A', 'R'), headersList):
        worksheetAgingP1.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('S6:Z6', 'PAST DUE', workbookFormat(workbook, headerStyle))

    for x, y in zip(alphabetRange('S', 'Z'), headersList1):
        worksheetAgingP1.write('{}7'.format(x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('AA6:AA7', 'TOTAL DUE', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.merge_range('AB6:AE6', 'PAST DUE BREAKDOWN', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AB7', 'PRINCIPAL', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AC7', 'INTEREST', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AD7', 'PENALTY', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AE7', 'TOTAL', workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('A{}:AE{}'.format(count1, count1), agingp1nodisplay, workbookFormat(workbook, entriesStyle))
    worksheetAgingP1.merge_range('A{}:C{}'.format(count1 + 1, count1 + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    worksheetAgingP1.write('M{}'.format(count1 + 1), "=SUM(M8:M{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    for c in range(ord('O'), ord('Z') + 1):
        if(chr(c) == 'Q'):
            worksheetAgingP1.write('Q{}'.format(count1 + 1), "=SUM(Q8:Q{})".format(count1 - 1),
                                   workbookFormat(workbook, sumStyle))
        else:
            worksheetAgingP1.write('{}{}'.format(chr(c), count1 + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count1 - 1),
                                workbookFormat(workbook, footerStyle))

    worksheetAgingP1.write('AA{}'.format(count1 + 1), "=SUM(AA8:AA{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AB{}'.format(count1 + 1), "=SUM(AB8:AB{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AC{}'.format(count1 + 1), "=SUM(AC8:AC{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AD{}'.format(count1 + 1), "=SUM(AD8:AD{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AE{}'.format(count1 + 1), "=SUM(AE8:AE{})".format(count1 - 1), workbookFormat(workbook, footerStyle))


    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/operationAgingReport", methods=['GET'])
def operationAgingReport():

    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')

    payload = {'date': date}

    url = lambdaUrl.format("reports/operationAging")

    r = requests.get(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "ADDRESS", "TERM", "FDD", "STATUS",
               "PNV", "MLV", "bPNV", "bMLV", "MI", "NOT DUE", "MATURED", "TODAY", "1-30", "31-60", "61-90", "91-120",
               "121-150", "151-180", "181-360", "360 & OVER", "TOTAL", "DUE PRINCIPAL", "DUE INTEREST", "DUE PENALTY"]
    df = pd.DataFrame(data['operationAgingReportJson'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 28)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['loanaccountNumber'] = df['loanaccountNumber'].map(lambda x: x.lstrip("'"))
        df['fdd'] = pd.to_datetime(df['fdd'])
        df['fdd'] = df['fdd'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df = round(df, 2)
        df['num'] = numbers(df.shape[0])
        df = df[["num", "appId", "loanaccountNumber", "fullName", "mobile", "address", "term", "fdd", "status", "PNV",
                 "MLV", "bPNV", "bMLV", "mi", "notDue", "matured", "today", "1-30", "31-60", "61-90", "91-120",
                 "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
    df = df.style.set_properties(**styles)
    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book
    merge_format2 = workbook.add_format(docNameStyle)
    merge_format4 = workbook.add_format(footerStyle)
    merge_format6 = workbook.add_format(entriesStyle)
    merge_format7 = workbook.add_format(headerStyle)
    merge_format8 = workbook.add_format(sumStyle)

    worksheet = writer.sheets["Sheet_1"]

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    worksheet.freeze_panes(5, 0)

    range1 = 'Y'
    range2 = 'Z'
    range3 = 'AC'
    companyName = 'RFSC'
    reportTitle = 'Operation Aging Report'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range2), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), merge_format7)

    worksheet.merge_range('AA6:AA7', 'DUE PRINCIPAL', merge_format7)
    worksheet.merge_range('AB6:AB7', 'DUE INTEREST', merge_format7)
    worksheet.merge_range('AC6:AC7', 'DUE PENALTY', merge_format7)

    worksheet.merge_range('A{}:AC{}'.format(count, count), nodisplay, merge_format6)
    worksheet.merge_range('A{}:B{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)

    for c in range(ord('J'), ord('Z') + 1):
        if (chr(c) == 'P'):
            worksheet.write('P{}'.format(count + 1), "=SUM(P8:P{})".format(count - 1), merge_format8)
        else:
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                                merge_format4)

    worksheet.write('AA{}'.format(count + 1), "=SUM(AA8:AA{})".format(count - 1), merge_format4)
    worksheet.write('AB{}'.format(count + 1), "=SUM(AB8:AB{})".format(count - 1), merge_format4)
    worksheet.write('AC{}'.format(count + 1), "=SUM(AC8:AC{})".format(count - 1), merge_format4)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Operations) as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/newmemoreport", methods=['GET'])
def newmemoreport():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = lambdaUrl.format("reports/memoreport")
    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "SUB PRODUCT", "MEMO TYPE", "PURPOSE", "AMOUNT",
               "STATUS", "DATE", "CREATED BY", "REMARKS", "APPROVED DATE", "APPROVED BY", "APPROVED REAMARKS"]
    creditDf = pd.DataFrame(data['Credit'])

    list1 = [len(i) for i in headers]
    if creditDf.empty:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = 'No Data'
        creditDf = pd.DataFrame(pd.np.empty((0, 14)))
        creditlist2 = list1
    else:
        countCredit = creditDf.shape[0] + 8
        nodisplayCredit = ''
        astype(creditDf, 'appId', int)
        creditDf.sort_values(by=['appId'], inplace=True)
        creditDf['loanAccountNo'] = creditDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        # creditDf['date'] = creditDf.date.apply(lambda x: x.split(" ")[0])
        dfDateFormat(creditDf, 'approvedDate')
        dfDateFormat(creditDf, 'date')
        creditDf['num'] = numbers(creditDf.shape[0])
        creditDf = creditDf[["num", "appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]
        creditlist2 = [max([len(str(s)) for s in creditDf[col].values]) for col in creditDf.columns]

    debitDf = pd.DataFrame(data['Debit'])

    if debitDf.empty:
        debitDf = pd.DataFrame(pd.np.empty((0, 14)))
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = 'No Data'
        debitlist2 = list1
    else:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = ''
        astype(debitDf, 'appId', int)
        debitDf.sort_values(by=['appId'], inplace=True)
        debitDf['loanAccountNo'] = debitDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        # debitDf['date'] = creditDf.date.apply(lambda x: x.split(" ")[0])
        dfDateFormat(debitDf, 'approvedDate')
        dfDateFormat(debitDf, 'date')
        debitDf['num'] = numbers(debitDf.shape[0])
        debitDf = debitDf[["num", "appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                           "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]
        debitlist2 = [max([len(str(s)) for s in debitDf[col].values]) for col in debitDf.columns]

    creditDf.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Credit", header=None)
    debitDf.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Debit", header=None)

    workbook = writer.book

    worksheetCredit = writer.sheets["Credit"]

    dataframeStyle(worksheetCredit, 'A', 'B', 8, countCredit, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetCredit, 'C', 'G', 8, countCredit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetCredit, 'H', 'H', 8, countCredit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetCredit, 'I', 'I', 8, countCredit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetCredit, 'J', 'J', 8, countCredit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetCredit, 'K', 'L', 8, countCredit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetCredit, 'M', 'M', 8, countCredit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetCredit, 'N', 'O', 8, countCredit, workbookFormat(workbook, stringFormat))

    for col_num, value in enumerate(columnWidth(list1, creditlist2)):
        worksheetCredit.set_column(col_num, col_num, value)

    range1 = 'L'
    range2 = 'M'
    range3 = 'O'
    companyName = 'RFSC'
    creditReportTitle = 'Memo Report (Credit)'
    debitReportTitle = 'Memo Report (Debit)'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(dateStart, dateEnd)

    workSheet(workbook, worksheetCredit, range1, range2, range3, xldate_header, name, companyName, creditReportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheetCredit.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetCredit.merge_range('A{}:O{}'.format(countCredit, countCredit), nodisplayCredit, workbookFormat(workbook, entriesStyle))
    worksheetCredit.merge_range('A{}:C{}'.format(countCredit + 1, countCredit + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheetCredit.write('H{}'.format(countCredit + 1), "=SUM(H8:H{})".format(countCredit - 1), workbookFormat(workbook, footerStyle))

    worksheetDebit = writer.sheets["Debit"]

    dataframeStyle(worksheetDebit, 'A', 'B', 8, countDebit, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetDebit, 'C', 'G', 8, countDebit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetDebit, 'H', 'H', 8, countDebit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetDebit, 'I', 'I', 8, countDebit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetDebit, 'J', 'J', 8, countDebit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetDebit, 'K', 'L', 8, countDebit, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetDebit, 'M', 'M', 8, countDebit, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetDebit, 'N', 'O', 8, countDebit, workbookFormat(workbook, stringFormat))
    for col_num, value in enumerate(columnWidth(list1, debitlist2)):
        worksheetDebit.set_column(col_num, col_num, value)

    workSheet(workbook, worksheetDebit, range1, range2, range3, xldate_header, name, companyName, debitReportTitle, branchName)

    headersList2 = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList2):
        worksheetDebit.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetDebit.merge_range('A{}:O{}'.format(countDebit, countDebit), nodisplayDebit, workbookFormat(workbook, entriesStyle))
    worksheetDebit.merge_range('A{}:C{}'.format(countDebit + 1, countDebit + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheetDebit.write('H{}'.format(countDebit + 1), "=SUM(H8:H{})".format(countDebit - 1), workbookFormat(workbook, footerStyle))

    writer.close()

    output.seek(0)
    print('sending spreadsheet')
    filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/tat", methods=['GET'])
def tat():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = lambdaUrl.format("newtat")

    r = requests.post(url, json=payload)
    data = r.json()
    standard = data['standard']
    returned = data['return']

    standardHeaders = [
         "#", "APP ID", "FIRST NAME", "LAST NAME", "MLV", "PNV", "APP DATE", "APP TIME", "PRODUCT", "STATUS",
         "PENDING - FOR VERIFICATION", "FOR VERIFICATION - FOR ADJUDICATION", "FOR VERIFICATION - FOR CANCELLATION", "FOR CANCELLATION - CANCELLED",
         "FOR ADJUDICATION - FOR APPROVAL", "FOR APPROVAL - APPROVED", "FOR APPROVAL - DISAPPROVED", "APPROVED - FOR RELEASING",
         "FOR RELEASING - RELEASED"]

    returnedHeaders = [
         "#", "APP ID", "FIRST NAME", "LAST NAME", "MLV", "PNV", "APP DATE", "APP TIME", "PRODUCT", "STATUS",
         "PENDING - FOR VERIFICATION", "FOR VERIFICATION - FOR ADJUDICATION", "FOR VERIFICATION - FOR CANCELLATION", "FOR CANCELLATION - CANCELLED",
         "FOR ADJUDICATION - REVERIFY", "REVERIFY - FOR ADJUDICATION", "FOR ADJUDICATION - FOR APPROVAL", "FOR APPROVAL - REVERIFY",
         "FOR APPROVAL - READJUDICATE", "READJUDICATE - FOR APPROVAL", "FOR APPROVAL - APPROVED", "FOR APPROVAL - DISAPPROVED",
         "APPROVED - FOR RELEASING", "FOR RELEASING - RELEASED"]

    standardlist1 = [len(i) for i in standardHeaders]
    returnedlist1 = [len(i) for i in returnedHeaders]

    standard_df = pd.read_csv(StringIO(standard))
    returned_df = pd.read_csv(StringIO(returned))

    countStandard = standard_df.shape[0] + 8
    countReturned = returned_df.shape[0] + 8

    dfDateFormat(standard_df, 'Application Date')
    dfDateFormat(returned_df, 'Application Date')

    standard_df.insert(0, column='#', value=numbers(standard_df.shape[0]))
    returned_df.insert(0, column='#', value=numbers(returned_df.shape[0]))

    if standard_df.empty:
        nodisplayStandard = 'No Data'
        standardlist2 = standardlist1

    else:
        nodisplayStandard = ''
        standardlist2 = [max([len(str(s)) for s in standard_df[col].values]) for col in standard_df.columns]

    if returned_df.empty:
        nodisplayReturned = 'No Data'
        returnedlist2 = returnedlist1
    else:
        nodisplayReturned = ''
        returnedlist2 = [max([len(str(s)) for s in returned_df[col].values]) for col in returned_df.columns]

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    standard_df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Standard", header=None)
    returned_df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Returned", header=None)

    workbook = writer.book

    worksheetStandard = writer.sheets["Standard"]

    dataframeStyle(worksheetStandard, 'A', 'B', 8, countStandard, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetStandard, 'C', 'D', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetStandard, 'E', 'H', 8, countStandard, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetStandard, 'I', 'J', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetStandard, 'K', 'S', 8, countStandard, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(standardlist1, standardlist2)):
        worksheetStandard.set_column(col_num, col_num, value)

    range1 = 'P'
    range2 = 'Q'
    range3 = 'S'
    companyName = 'RFSC'
    reportTitle = 'TAT Report (Standard)'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheetStandard, range1, range2, range3, xldate_header, name, companyName, reportTitle,
              branchName)

    headersListstandard = [i for i in standardHeaders]

    for x, y in zip(alphabet(range3), headersListstandard):
        worksheetStandard.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetStandard.merge_range('A{}:S{}'.format(countStandard, countStandard), nodisplayStandard, workbookFormat(workbook, entriesStyle))

    for c in range(ord('K'), ord('S') + 1):
        worksheetStandard.write('{}{}'.format(chr(c), countStandard + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), countStandard - 1),
                        workbookFormat(workbook, sumStyle))

    worksheetReturned = writer.sheets["Returned"]

    dataframeStyle(worksheetReturned, 'A', 'B', 8, countStandard, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetReturned, 'C', 'D', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetReturned, 'E', 'H', 8, countStandard, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetReturned, 'I', 'J', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetReturned, 'K', 'X', 8, countStandard, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(returnedlist1, returnedlist2)):
        worksheetReturned.set_column(col_num, col_num, value)

    range1 = 'U'
    range2 = 'V'
    range3 = 'X'
    companyName = 'RFSC'
    reportTitle = 'TAT Report (Returned)'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheetReturned, range1, range2, range3, xldate_header, name, companyName, reportTitle,
              branchName)

    headersListreturned = [i for i in returnedHeaders]

    for x, y in zip(alphabet(range3), headersListreturned):
        worksheetReturned.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetReturned.merge_range('A{}:X{}'.format(countReturned, countReturned), nodisplayReturned, workbookFormat(workbook, entriesStyle))

    for c in range(ord('K'), ord('X') + 1):
        worksheetReturned.write('{}{}'.format(chr(c), countReturned + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), countReturned - 1),
                        workbookFormat(workbook, sumStyle))

    writer.close()
    output.seek(0)

    filename = "TAT {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/unappliedbalances", methods=['GET'])
def get_uabalances():
    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')
    payload = {}

    url = serviceUrl.format("accountDueReportJSON")

    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "AMOUNT DUE", "DUE DATE",
               "UNAPPLIED BALANCE"]
    df = pd.DataFrame(data['accountDueReportJSONResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 7)))
        list2 = list1
    else:
        nodisplay = ''
        count = df.shape[0] + 8
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['loanId'], inplace=True)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['dueDate'] = pd.to_datetime(df['dueDate'])
        df['dueDate'] = df['dueDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'dueDate')
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'E', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'F', 'H', 8, count, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'E'
    range2 = 'F'
    range3 = 'H'
    companyName = 'RFSC'
    reportTitle = 'Unapplied Balance Report'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:H{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    worksheet.write('F{}'.format(count + 1), "=SUM(F8:F{})".format(count - 1), workbookFormat(workbook, footerStyle))
    worksheet.write('H{}'.format(count + 1), "=SUM(H8:H{})".format(count - 1), workbookFormat(workbook, footerStyle))
    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Unapplied Balance {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/dccr", methods=['GET'])
def get_data1():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    print('US/Pacific', now_pacific)
    print('generation date', dateNow)

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("DCCRjsonNew")

    r = requests.post(url, json=payload)
    data_json = r.json()

    # sortData = sorted(data_json['DCCRjsonNewResult'], key=lambda d: d['orNo'], reverse=False)
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers1 = ["#", "PAYMENT", "LOAN ACCT. #", "CUSTOMER NAME", "OR DATE", "OR NUM", "BANK", "CHECK #", "PAYMENT",
                "TOTAL", "CASH", "CHECK", "PRINCIPAL", "INTEREST", "ADVANCES", "PENALTY"]
    headers = ["LOAN ACCT. #", "CUSTOMER NAME", "OR DATE", "OR #", "BANK", "CHECK #"]
    df = pd.DataFrame(data_json['DCCRjsonNewResult'])
    df1 = pd.DataFrame(data_json['DCCRjsonNewResult']).copy()
    list1 = [len(i) for i in headers1]
    if df.empty or df1.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 25)))
        dfCashcount = 0
        dfEcpaycount = 0
        dfBCcount = 0
        dfBankcount = 0
        dfCheckcount = 0
        dfGPRScount = 0
        df1['num1'] = ''
        dfCash = pd.DataFrame(pd.np.empty((0, 25)))
        dfEcpay = pd.DataFrame(pd.np.empty((0, 25)))
        dfBC = pd.DataFrame(pd.np.empty((0, 25)))
        dfBank = pd.DataFrame(pd.np.empty((0, 25)))
        dfCheck = pd.DataFrame(pd.np.empty((0, 25)))
        dfGPRS = pd.DataFrame(pd.np.empty((0, 25)))
        df2 = pd.DataFrame(pd.np.empty((0, 25)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        astype(df, 'orNo', int)
        df.sort_values(by=['orNo'], inplace=True)
        conditions = [(df['transType'] == 'Check')]
        conditionBank = [(df['transType'] == 'Bank')]
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['total'] = np.select(conditions, [df['paymentCheck']], default=df['amount'])
        df['total1'] = np.select(conditions, [df['paymentCheck']], default=df['amount'])
        df1['total'] = np.select(conditions, [df1['amount']], default=0)
        df['paymentCheck'] = np.select(conditions, [df['amount']], default=0)
        df['cash'] = np.select(conditions,[0], default=df['amount'])
        df['date'] = np.select(conditions, [df['checkDate']], default=df['paymentDate'])
        df['check'] = np.select(conditions, [df['paymentSource']], default='')
        df['bank'] = np.select(conditionBank, [df['paymentSource']], default=df['check'])
        diff = df['amount'] - (df['paidPrincipal'] + df['paidInterest'] + df['paidPenalty'])
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        dfDateFormat(df, 'orDate')
        dfDateFormat(df, 'date')
        df['advances'] = round(diff, 2)
        df['num'] = numbers(df.shape[0])
        df1['num1'] = ''
        df['num1'] = ''
        df = round(df, 2)
        df1 = round(df, 2)
        df1 = df1.sort_values(by=['transType'])
        dfCash = df1.loc[df1['transType'] == 'Cash'].copy()
        dfEcpay = df1.loc[df1['transType'] == 'Ecpay'].copy()
        dfBC = df1.loc[df1['transType'] == 'Bayad Center'].copy()
        dfCheck = df1.loc[df1['transType'] == 'Check'].copy()
        dfBank = df1.loc[df1['transType'] == 'Bank'].copy()
        dfGPRS = df1.loc[df1['transType'] == 'GPRS'].copy()
        astype(dfCash, 'orNo', int)
        astype(dfEcpay, 'orNo', int)
        astype(dfBC, 'orNo', int)
        astype(dfCheck, 'orNo', int)
        astype(dfBank, 'orNo', int)
        astype(dfGPRS, 'orNo', int)
        dfCash.sort_values(by=['orNo'], inplace=True)
        dfEcpay.sort_values(by=['orNo'], inplace=True)
        dfBC.sort_values(by=['orNo'], inplace=True)
        dfCheck.sort_values(by=['orNo'], inplace=True)
        dfBank.sort_values(by=['orNo'], inplace=True)
        dfGPRS.sort_values(by=['orNo'], inplace=True)
        dfCashcount = dfCash.shape[0]
        dfEcpaycount = dfEcpay.shape[0]
        dfBCcount = dfBC.shape[0]
        dfBankcount = dfBank.shape[0]
        dfCheckcount = dfCheck.shape[0]
        dfGPRScount = dfGPRS.shape[0]

        dfCash['dfCashnum'] = numbers(dfCashcount)
        dfEcpay['dfEcpaynum'] = numbers(dfEcpaycount)
        dfBC['dfBCnum'] = numbers(dfBCcount)
        dfBank['dfBanknum'] = numbers(dfBankcount)
        dfCheck['dfChecknum'] = numbers(dfCheckcount)
        dfGPRS['dfGPRSnum'] = numbers(dfGPRScount)
        df = df[['num', 'transType', 'loanAccountNo', 'newCustomerName', 'orDate', 'orNo', 'bank', 'checkNo', 'date',
                 'amount', 'cash', 'paymentCheck', 'paidPrincipal', 'paidInterest', 'advances', 'paidPenalty']]
        dfCash = dfCash[['dfCashnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfEcpay = dfEcpay[['dfEcpaynum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBC = dfBC[['dfBCnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBank = dfBank[['dfBanknum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfCheck = dfCheck[['dfChecknum', 'orDate', 'orNo', 'transType', 'amount', 'total', 'paymentCheck']]
        dfGPRS = dfGPRS[['dfGPRSnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        df2 = df1[['num1', 'num1', 'num1', 'num1']]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    workbook = writer.book

    worksheet = workbook.add_worksheet('Sheet_1')
    writer.sheets['Sheet_1'] = worksheet
    # dfEcpay = dfEcpay.style.set_properties(**styles)
    # dfBC = dfBC.style.set_properties(**styles)
    # dfBank = dfBank.style.set_properties(**styles)
    # dfCash = dfCash.style.set_properties(**styles)
    # df = df.style.set_properties(**styles)
    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    if(dfCashcount <= 0):
        dfwriter(dfEcpay.to_excel, writer, count + 10)
        dfwriter(dfBC.to_excel, writer, count + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfEcpaycount + dfBCcount + dfBankcount + 25)
        dfwriter(dfGPRS.to_excel, writer, count + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + 30)
    elif (dfEcpaycount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfBCcount + dfBankcount + 25)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfBCcount + dfBankcount + dfCheckcount + 30)
    elif (dfBCcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBankcount + 25)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBankcount + dfCheckcount + 30)
    elif (dfBankcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfCheckcount + 25)
    elif (dfCheckcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 25)
    elif (dfGPRScount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 25)
    else:
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 25)
        dfwriter(dfGPRS.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + 30)

    dfwriter(df2.to_excel, writer, count + count + 2)

    dataframeStyle(worksheet, 'A', 'A', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'B', 'D', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'E', 'F', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'G', 'G', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'H', 'I', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'J', 'P', 8, count, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        if(col_num == 2):
            worksheet.set_column(2, 2, 13)
        elif (col_num == 3):
            worksheet.set_column(3, 3, 23)
        else:
            worksheet.set_column(col_num, col_num, value)

    # worksheet.freeze_panes(5, 0)

    range1 = 'I'
    range2 = 'L'
    range3 = 'P'
    companyName = 'RFSC'
    reportTitle = 'DAILY CASH/CHECK COLLECTION REPORT'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    def alphabetRange(firstRange, secondRange):
        alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
        return alphaList

    for x, y in zip(alphabetRange('C', 'I'), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A6:A7', '#', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('B6:B7', 'PAYMENT\nCHANNEL', workbookFormat(workbook, textWrapHeader))
    worksheet.merge_range('I6:I7', 'PAYMENT\nDATE', workbookFormat(workbook, textWrapHeader))
    worksheet.write('J7', 'TOTAL', workbookFormat(workbook, headerStyle))
    worksheet.write('K7', 'CASH', workbookFormat(workbook, headerStyle))
    worksheet.write('L7', 'CHECK', workbookFormat(workbook, headerStyle))
    worksheet.write('M7', 'PRINCIPAL', workbookFormat(workbook, headerStyle))
    worksheet.write('N7', 'INTEREST', workbookFormat(workbook, headerStyle))
    worksheet.write('O7', 'ADVANCES', workbookFormat(workbook, headerStyle))
    worksheet.write('P7', 'PENALTY\n(5%)', workbookFormat(workbook, textWrapHeader))

    worksheet.merge_range('J6:L6', 'AMOUNT', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('M6:P6', 'LOAN REPAYMENT', workbookFormat(workbook, headerStyle))

    worksheet.merge_range('I{}:P{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.write('I{}'.format(count + 1), 'TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheet.merge_range('A{}:P{}'.format(count + 2, count + 2), '', workbookFormat(workbook, topBorderStyle))

    for c in range(ord('J'), ord('P') + 1):
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                            workbookFormat(workbook, footerStyle))

    countcash = count + dfCashcount
    countecpay = count + dfCashcount + dfEcpaycount + 5
    countbc = count + dfCashcount + dfEcpaycount + dfBCcount + 10
    countbank = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 15
    countcheck = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + 20
    countgprs = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + dfGPRScount + 25

    paymentTypeWorksheet(worksheet, count, 'CASH TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countcash + 5, 'ECPAY TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countecpay + 5, 'BAYAD CENTER\nTYPE', workbookFormat(workbook, textWrapHeader))
    paymentTypeWorksheet(worksheet, countbc + 5, 'BANK TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countbank + 5, 'CHECK TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countcheck + 5, 'GPRS TYPE', workbookFormat(workbook, headerStyle))

    dataframeStyle(worksheet, 'E', 'E', count + 6, countcash + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countcash + 11, count + countecpay + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countecpay + 11, countbc + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countbc + 6, count + countbank + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countbank + 11, countcheck + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countcheck + 11, countgprs + 5, workbookFormat(workbook, numFormat))

    sumPaymentType(worksheet, countcash, count, countcash, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countecpay, countcash + 5, countecpay, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countbc, countecpay + 5, countbc, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countbank, countbc + 5, countbank, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countcheck, countbank + 5, countcheck, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countgprs, countcheck + 5, countgprs, workbookFormat(workbook, footerStyle))

    totalPaymentType(worksheet, countcash, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countecpay, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countbc, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countbank, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countcheck, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countgprs, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))

    counts = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + dfGPRScount + 10
    worksheet.merge_range('A{}:D{}'.format(counts + 24, counts + 24), 'DISBURSMENT', workbookFormat(workbook, headerStyle))
    worksheet.write('A{}'.format(counts + 25), '#', workbookFormat(workbook, headerStyle))
    worksheet.write('B{}'.format(counts + 25), 'DATE', workbookFormat(workbook, headerStyle))
    worksheet.write('C{}'.format(counts + 25), 'DESCRIPTION', workbookFormat(workbook, headerStyle))
    worksheet.write('D{}'.format(counts + 25), 'AMOUNT', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('A{}:B{}'.format(counts + 27, counts + 27), 'TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheet.write('D{}'.format(counts + 27), "=SUM(D{}:D{})".format(counts + 26, counts + 26), workbookFormat(workbook, borderFormatStyle))
    worksheet.merge_range('A{}:C{}'.format(counts + 29, counts + 29), 'NET COLLECTION:', workbookFormat(workbook, docNameStyle))
    worksheet.write('D{}'.format(counts + 29), "=J{}-D{}".format(count + 1, counts + 27), workbookFormat(workbook, borderFormatStyle))
    # worksheet.write('C{}'.format(count + count + 1), nodisplay, merge_format8)

    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/monthlyincome", methods=['GET'])
def get_monthly1():

    output = BytesIO()

    date = request.args.get('date')
    name = request.args.get('name')
    datetime_object = datetime.strptime(date, '%m/%d/%Y')
    month = datetime_object.strftime("%B")

    payload = {'date': date}

    url = serviceUrl.format("monthlyIncomeReportJs")

    r = requests.post(url, json=payload)
    data_json = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "PENALTY PAID",
               "INTEREST PAID", "PRINCIPAL PAID", "UNAPPLIED BALANCE", "PAYMENT AMOUNT", "OR DATE", "OR #"]
    df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 10)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['loanAccountno'] = df['loanAccountno'].map(lambda x: x.lstrip("'"))
        astype(df, 'appId', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['appId', 'orDate'], inplace=True)
        # df['orAmount'] = 0
        # df["unappliedBalance"] = (df['penaltyPaid'] + df['interestPaid'] + df['principalPaid']) - df['orAmount']
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'orDate')
        df = round(df, 2)
        df = df[['num', 'appId', 'loanAccountno', 'newCustomerName', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
                 'paymentAmount', "orDate", "orNo"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'D', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'E', 'J', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'K', 'K', 8, count, workbookFormat(workbook, stringFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'H'
    range2 = 'I'
    range3 = 'K'
    companyName = 'RFSC'
    reportTitle = 'Mothly Income Report'
    branchName = 'Nationwide'
    xldate_header = "For the month of {}".format(month)

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:K{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    for c in range(ord('E'), ord('I') + 1):
        worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                        workbookFormat(workbook, footerStyle))

    writer.close()

    output.seek(0)
    print('sending spreadsheet')
    filename = "Monthly Income {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/booking", methods=['GET'])
def get_booking():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("bookingReportJs")

    r = requests.post(url, json=payload)
    data_json = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "PRODUCT CODE", "SA", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "SUB PRODUCT", "PNV", "MLV", "FINANCE FEE", "HF",
               "DST", "NF", "GCLI", "OMA", "TERM (MOS)", "RATE(%)", "MI", "APPLICATION DATE", "APPROVAL DATE", "BOOKING DATE", "FDD", "PROMO NAME"]
    df = pd.DataFrame(data_json['bookingReportJsResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 23)))
        list2 = list1
    else:
        nodisplay = ''
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['forreleasingdate'] = df.forreleasingdate.apply(lambda x: x.split(" ")[0])
        df['approvalDate'] = df.approvalDate.apply(lambda x: x.split(" ")[0])
        df['applicationDate'] = df.applicationDate.apply(lambda x: x.split(" ")[0])
        df['generationDate'] = df.generationDate.apply(lambda x: x.split(" ")[0])
        df["customerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df["actualRate"] = df["actualRate"] + "%"
        dfDateFormat(df, 'forreleasingdate')
        dfDateFormat(df, 'approvalDate')
        dfDateFormat(df, 'generationDate')
        dfDateFormat(df, 'applicationDate')
        dfDateFormat(df, 'fdd')
        astype(df, 'loanId', int)
        astype(df, 'term', int)
        # astype(df, 'actualRate', float)
        df.sort_values(by=['loanId', 'forreleasingdate'], inplace=True)
        count = df.shape[0] + 8
        df['num'] = numbers(df.shape[0])
        df = df[['num', 'channelName', 'partnerCode', 'outletCode', 'productCode', 'sa', 'loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "mlv", "interest",
                 "handlingFee", "dst", "notarial", "gcli", "otherFees", "term", "actualRate", "monthlyAmount", 'applicationDate', 'approvalDate', 'forreleasingdate', 'fdd',
                 'promoName']]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    # df = df.style.set_properties(**styles)
    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'A', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'B', 'F', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'G', 'G', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'H', 'J', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'K', 'R', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'S', 'S', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'T', 'X', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'Y', 'Z', 8, count, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    # for col_num, value in enumerate(alphabet('K')):
    #     worksheet.merge_range('{}8:{}{}'.format(value, value, count - 1), merge_format8)

    # worksheet.freeze_panes(7, 0)

    range1 = 'W'
    range2 = 'X'
    range3 = 'Z'
    companyName = 'RFSC'
    reportTitle = 'Booking Report'
    branchName = 'Nationwide'

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:Z{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:B{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    for c in range(ord('K'), ord('R') + 1):
        worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                        workbookFormat(workbook, footerStyle))
    worksheet.write('U{}'.format(count + 1), "=SUM(U8:U{})".format(count - 1), workbookFormat(workbook, footerStyle))

    writer.close()

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

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("generateincentiveReportJSON")

    r = requests.post(url, json=payload)
    data_json = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "BOOKING DATE", "APP ID", "CLIENT'S NAME", "REFERRAL TYPE", "SA", "BRANCH", "LOAN TYPE",  "TERM", "MLV", "PNV",
               "MI", "REFERRER"]
    df = pd.DataFrame(data_json['generateincentiveReportJSONResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 12)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['agentName'], inplace=True)
        df['bookingDate'] = pd.to_datetime(df['bookingDate'])
        df['bookingDate'] = df['bookingDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'bookingDate')
        df = df[['num', 'bookingDate', 'loanId', 'newCustomerName', 'refferalType', "SA", "dealerName", "loanType", "term",
             "totalAmount", "PNV", "monthlyAmount", "agentName"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'C', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'D', 'H', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'I', 'I', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'J', 'L', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'M', 'M', 8, count, workbookFormat(workbook, stringFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'J'
    range2 = 'K'
    range3 = 'M'
    companyName = 'RFSC'
    reportTitle = 'Sales Referral Report'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:M{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    for c in range(ord('J'), ord('L') + 1):
        worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                        workbookFormat(workbook, footerStyle))

    writer.close()

    output.seek(0)
    print('sending spreadsheet')
    filename = "Sales Referral Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/mature", methods=['GET'])
def get_mature():

    output = BytesIO()

    date = request.args.get('date')
    name = request.args.get('name')

    payload = {'date': date}

    url = serviceUrl.format("maturedLoanReport")

    r = requests.post(url, json=payload)
    data_json = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "TERM", "BMLV", "LAST DUE DATE",
               "LAST PAYMENT", "NO. OF UNPAID", "TOTAL PAYMENT", "TOTAL PAST DUE", "OB",
               "NO. OF MONTHS"]
    df = pd.DataFrame(data_json['maturedLoanReportResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 13)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        astype(df, 'monthlydue', float)
        astype(df, 'outStandingBalance', float)
        astype(df, 'loanId', int)
        astype(df, 'unpaidMonths', int)
        astype(df, 'term', int)
        astype(df, 'matured', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        dfDateFormat(df, 'lastDueDate')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        df = df[['num', 'loanId', 'loanAccountNo', 'newCustomerName', "mobileno", "term", "bMLV", "lastDueDate", "lastPayment",
                 "unpaidMonths", "totalPayment", "monthlydue", "outStandingBalance", "matured"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'E', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'F', 'F', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'G', 'I', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'J', 'J', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'K', 'M', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'N', 'N', 8, count, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'K'
    range2 = 'L'
    range3 = 'N'
    companyName = 'RFSC'
    reportTitle = 'Matured Loans Report'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        if(x == 'N'):
            worksheet.merge_range('N6:N7', 'NO. OF MONTHS\nFROM MATURITY', workbookFormat(workbook, textWrapHeader))
        elif(x == 'J'):
            worksheet.merge_range('J6:J7', 'NO. OF UNPAID\nMONTHS', workbookFormat(workbook, textWrapHeader))
        else:
            worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:N{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    worksheet.write('G{}'.format(count + 1), "=SUM(G8:G{})".format(count - 1), workbookFormat(workbook, footerStyle))
    for c in range(ord('K'), ord('M') + 1):
        worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                        workbookFormat(workbook, footerStyle))

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

    payload = {'date': date}

    url = serviceUrl.format("dueTodayReport")

    r = requests.post(url, json=payload)
    data_json = r.json()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "LOAN TYPE", "DUE TODAY TERM",
               "MI", "TOTAL PAST DUE", "UNPAID PENALTY", "MONTHLY DUE", "LAST PAYMENT DATE", "LAST PAYMENT AMOUNT"]
    df = pd.DataFrame(data_json['dueTodayReportResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 12)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        astype(df, 'monthlyAmmortization', float)
        astype(df, 'monthdue', float)
        astype(df, 'loanId', int)
        astype(df, 'term', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        dfDateFormat(df, 'monthlydue')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileno", "loanType", "term", "monthlyAmmortization",
             "monthdue", "unpaidPenalty", "monthlydue", "lastPayment", "lastPaymentAmount"]]
        list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]

    df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'F', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'G', 'G', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'H', 'M', 8, count, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'J'
    range2 = 'K'
    range3 = 'M'
    companyName = 'RFSC'
    reportTitle = 'Due Today Report'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:M{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))
    for c in range(ord('H'), ord('J') + 1):
        worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                        workbookFormat(workbook, footerStyle))

    worksheet.write('M{}'.format(count + 1), "=SUM(M8:M{})".format(count - 1), workbookFormat(workbook, footerStyle))

    writer.close()

    output.seek(0)
    print('sending spreadsheet')
    filename = "Due Today Report {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/customerLedger", methods=['GET'])
def get_customerLedger():

    output = BytesIO()

    loanId = request.args.get('loanId')
    userId = request.args.get('userId')
    date = request.args.get('date')
    name = request.args.get('name')

    payload = {'loanId': loanId, 'userId': userId, 'date': date}

    ledgerById = "customerLedger/ledgerByLoanId?loanId={}".format(loanId)
    url = bluemixUrl.format(ledgerById)
    url2 = serviceUrl.format("getCustomerLedger")

    r = requests.post(url2, json=payload)
    data_json = r.json()
    ledgerData = requests.get(url).json()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    dfLedger = pd.DataFrame(ledgerData['data']['transactions'])
    dfCustomerLedger = pd.DataFrame(data_json['getCustomerLedgerResult'])

    headers = ["#", "DATE", "TERM", "TRANSACTION TYPE", "PAYMENT TYPE", "REF NO", "CHECK #", "PENALTY INCUR", "PRINCIPAL",
               "INTEREST", "PENALTY PAID", "ADVANCES", "TOTAL", "DUE", "OB", "PAYMENT DATE", "OR NO", "OR DATE"]

    headersBorrower = ["APPLICATION ID", "LOAN ACCOUNT NO.", "BORROWER'S NAME", "COLLECTOR", "CONTACT NO."]
    dataBorrower = ["appId", "loanAccountNo", "borrowersName", "collector", "contactNo", "address"]
    headersLoan = ["APPROVED LOAN AMOUNT", "LOAN TYPE", "GROSS MI", "TOTAL ADD-ON RATE", "TERMS", "DISBURSEMENT DATE", "FIRST DUE DATE"]
    dataLoan = ["loanAmount", "loanType", "mi", "addOnRate", "terms", "disbursementDate", "fdd"]
    headersCollateral = ["UNIT/MODEL/DESC.", "BRAND/MAKE", "SERIAL/CHASSIS NO."]
    dataCollateral = ["model", "brand", "serialNo"]
    headersAccStat = ["EXPIRED TERM", "REMAINING TERM", "NO. OF MI's PAID", "MONTHS DUE", "OVERDUE AMOUNT"]
    dataAcctStat = ["expiredTerm", "remainingTerm", "miPaid", "monthsDue", "overDueAmount", "lastPaymentDate"]
    headersOB = ["RFC", "PENALTY", "", "TOTAL", "", "TOTAL PAYMENT", "LAST PAYMENT DATE"]
    headersLoanSummary = ["TOTAL", "PAID", "ADJ", "BILLED", "AMT DUE", "BAL."]

    dfSummary = dfCustomerLedger['accountSummary']

    adjPrincipal = dfSummary['debitPrincipal'] + (dfSummary['creditPrincipal'] * -1)
    adjInterest = dfSummary['debitInterest'] + (dfSummary['creditInterest'] * -1)
    adjPenalty = dfSummary['debitPenalty'] + (dfSummary['creditPenalty'] * -1)

    # principalconditions = [(adjPrincipal < 0)]
    # prinPaid = np.select(principalconditions, [dfSummary['principalPaid'] - (adjPrincipal * -1)], default=dfSummary['principalPaid'] - adjPrincipal)
    #
    # interestconditions = [(adjInterest < 0)]
    # intPaid = np.select(interestconditions, [dfSummary['interestPaid'] - (adjInterest * -1)], default=dfSummary['interestPaid'] - adjInterest)
    #
    # penaltyconditions = [(adjPenalty < 0)]
    # penPaid = np.select(penaltyconditions, [dfSummary['penaltyPaid'] - (adjPenalty * -1)], default=dfSummary['penaltyPaid'] - adjPenalty)

    prinPaid = dfSummary['principalPaid'] - dfSummary['creditPrincipal']
    intPaid = dfSummary['interestPaid'] - dfSummary['creditInterest']
    penPaid = dfSummary['penaltyPaid'] - dfSummary['creditPenalty']

    prinTotal = dfSummary['principal'] - dfSummary['debitPrincipal']
    intTotal = dfSummary['interest'] - dfSummary['debitInterest']
    penTotal = dfSummary['penalty'] - dfSummary['debitPenalty']

    # dfSummary['principalPaid'] = dfSummary['principalPaid'] - dfSummary['principalAdj']
    # dfSummary['interestPaid'] = dfSummary['interestPaid'] - dfSummary['interestAdj']
    # dfSummary['penaltyPaid'] = dfSummary['penaltyPaid'] - dfSummary['penaltyAdj']

    loanPricipalData= ["principal", "principalPaid", "principalAdj", "principalBilled", "principalAmtDue"]
    loanInterestData= ["interest", "interestPaid", "interestAdj", "interestBilled", "interestAmtDue"]
    loanPenaltyData= ["penalty", "penaltyPaid", "penaltyAdj", "penaltyBilled", "penaltyAmtDue"]

    list1 = [len(i) for i in headers]

    if dfLedger.empty:
        count = dfLedger.shape[0] + 33
        nodisplay = 'No Data'
        dfLedger = pd.DataFrame(pd.np.empty((0, 12)))
        list2 = list1
    else:
        count = dfLedger.shape[0] + 33
        nodisplay = ''
        dfLedger['orDate'] = dfLedger['orDate'].loc[dfLedger['orDate'].str.contains("/")]
        dfLedger['paymentDate'] = dfLedger['paymentDate'].loc[dfLedger['paymentDate'].str.contains("/")]
        conditions = [(dfLedger['paymentDate'] == '-')]
        dfLedger['paymentDate'] = np.select(conditions, [dfLedger['paymentDate']], default="")
        dfDateFormat(dfLedger, 'orDate')
        dfDateFormat(dfLedger, 'paymentDate')
        dfDateFormat(dfLedger, 'date')
        astype(dfLedger, 'penaltyIncur', float)
        astype(dfLedger, 'principalPaid', float)
        astype(dfLedger, 'interestPaid', float)
        astype(dfLedger, 'penaltyPaid', float)
        astype(dfLedger, 'advances', float)
        astype(dfLedger, 'mi', float)
        astype(dfLedger, 'amountDue', float)
        astype(dfLedger, 'ob', float)
        dfLedger['total'] = dfLedger['penaltyIncur'] + dfLedger['principalPaid'] + dfLedger['interestPaid'] + dfLedger['penaltyPaid']
        dfLedger['num'] = numbers(dfLedger.shape[0])
        # paymentType = dfLedger.loc[dfLedger['type'] == 'Payment']
        # billed = dfLedger.loc[dfLedger['type'] == 'UPD']
        # totalInterestBilled= pd.Series(billed['interestPaid']).sum()
        # totalPrincipalBilled = pd.Series(billed['principalPaid']).sum()
        # totalPrincipalPaid = pd.Series(paymentType['principalPaid']).sum() * -1
        # totalInterestPaid = pd.Series(paymentType['interestPaid']).sum() * -1
        # totalPenalty = pd.Series(dfLedger['penaltyIncur']).sum()
        # totalPenaltyPaid = pd.Series(dfLedger['penaltyPaid']).sum() * -1
        dfLedger = round(dfLedger, 2)
        dfLedger = dfLedger[["num", "date", "term", "type", "paymentType", "refNo", "checkNo", "penaltyIncur", "principalPaid", "interestPaid", "penaltyPaid",
             "advances", "totalRow", "amountDue", "ob", "paymentDate", "orNo", "orDate"]]
        list2 = [max([len(str(s)) for s in dfLedger[col].values]) for col in dfLedger.columns]
    # print('totalPrincipalPaid', totalPrincipalPaid)
    # print('totalInterestPaid', totalInterestPaid)
    # print('totalPenalty', totalPenalty)
    dfCustomerLedger = round(dfCustomerLedger, 2)


    # dfLoanSummary = dfLoanSummary.style.set_properties(**styles)
    # dfLedger = dfLedger.style.set_properties(**styles)
    # dfLoanSummary.to_excel(writer, startrow=17, startcol=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)
    dfLedger.to_excel(writer, startrow=33, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'C', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'D', 'E', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'F', 'G', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'H', 'P', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'Q', 'R', 8, count, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'O'
    range2 = 'P'
    range3 = 'R'
    companyName = 'RFSC'
    reportTitle = 'CUSTOMER LEDGER'
    xldate_header = 'As of {}'.format(startDateFormat(date))
    branchName = 'Nationwide'

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    def cnumbers(numRange, addnum):
        number = [number + addnum for number in range(numRange)]
        return number

    def alphabetRange(firstRange, secondRange):
        alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
        return alphaList

    worksheet.merge_range('A6:F6', 'BORROWER DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(cnumbers(5, 7), headersBorrower):
        worksheet.merge_range('A{}:D{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    worksheet.merge_range('A12:D13', 'ADDRESS', workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(cnumbers(5, 7), dataBorrower):
        worksheet.merge_range('E{}:F{}'.format(x, x), dfCustomerLedger['borrower'][y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('E12:F13', dfCustomerLedger['borrower']['address'], workbookFormat(workbook, ledgerDataStyle2))

    worksheet.merge_range('H6:L6', 'LOAN DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(cnumbers(7, 7), headersLoan):
        worksheet.merge_range('H{}:J{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(cnumbers(7, 7), dataLoan):
        worksheet.merge_range('K{}:L{}'.format(x, x), dfCustomerLedger['loan'][y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('N6:R6', 'COLLATERAL DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(cnumbers(6, 7), headersCollateral):
        worksheet.merge_range('N{}:O{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(cnumbers(6, 7), dataCollateral):
        worksheet.merge_range('P{}:R{}'.format(x, x), dfCustomerLedger['collateral'][y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('A15:D15', 'ACCOUNT STATUS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(cnumbers(5, 16), headersAccStat):
        worksheet.merge_range('A{}:D{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    worksheet.merge_range('E15:F15', 'NORMAL', workbookFormat(workbook, undStyle))

    for x, y in zip(cnumbers(5, 16), dataAcctStat):
        worksheet.merge_range('E{}:F{}'.format(x, x), dfCustomerLedger['acctStat'][y], workbookFormat(workbook, defaultFormat))

    worksheet.merge_range('H15:O15', 'LOAN ACCOUNT SUMMARY', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(alphabetRange('J', 'O'), headersLoanSummary):
        worksheet.write('{}17'.format(x), '{}'.format(y), workbookFormat(workbook, sumStyle))

    for c in range(ord('J'), ord('N') + 1):
        if(chr(c) == 'J'):
            worksheet.write('J18', '=SUM(J22,J23,J25)', workbookFormat(workbook, ledgerNum))
        else:
            worksheet.write('{}18'.format(chr(c)), '=SUM({}22,{}23,{}25)'.format(chr(c),chr(c),chr(c)), workbookFormat(workbook, ledgerNum))

    for c in range(ord('J'), ord('N') + 1):
        worksheet.write('{}20'.format(chr(c)), '={}21'.format(chr(c)), workbookFormat(workbook, ledgerNum))

    for c in range(ord('J'), ord('N') + 1):
        worksheet.write('{}21'.format(chr(c)), '=SUM({}22,{}23)'.format(chr(c),chr(c)), workbookFormat(workbook, numFormat))

    # worksheet.write('H26', '=H18-H20', workbookFormat(workbook, numFormat))

    for num in range(18, 26):
        if( num == 19):
            worksheet.write('O19', None, workbookFormat(workbook, numFormat))
        elif(num == 24):
            worksheet.write('O24', None, workbookFormat(workbook, numFormat))
        elif(num == 18):
            worksheet.write('O18'.format(num), '=J18-K18+L18'.format(num, num, num), workbookFormat(workbook, ledgerNum))
        elif(num == 20):
            worksheet.write('O20'.format(num), '=J20-K20+L20'.format(num, num, num), workbookFormat(workbook, ledgerNum))
        else:
            worksheet.write('O{}'.format(num), '=J{}-K{}+L{}'.format(num,num,num), workbookFormat(workbook, numFormat))

    for x, y in zip(alphabetRange('J', 'N'), loanPricipalData):
        if (x == 'L'):
            worksheet.write('L22', adjPrincipal, workbookFormat(workbook, defaultFormat))
        elif (x == 'K'):
            worksheet.write('K22', prinPaid, workbookFormat(workbook, defaultFormat))
        elif (x == 'J'):
            worksheet.write('J22', prinTotal, workbookFormat(workbook, defaultFormat))
        else:
            worksheet.write('{}22'.format(x), dfCustomerLedger['accountSummary'][y],
                            workbookFormat(workbook, defaultFormat))

    for x, y in zip(alphabetRange('J', 'N'), loanInterestData):
        if (x == 'L'):
            worksheet.write('L23', adjInterest, workbookFormat(workbook, defaultFormat))
        elif (x == 'K'):
            worksheet.write('K23', intPaid, workbookFormat(workbook, defaultFormat))
        elif (x == 'J'):
            worksheet.write('J23', intTotal, workbookFormat(workbook, defaultFormat))
        else:
            worksheet.write('{}23'.format(x), dfCustomerLedger['accountSummary'][y],
                            workbookFormat(workbook, defaultFormat))

    for x, y in zip(alphabetRange('J', 'N'), loanPenaltyData):
        if (x == 'L'):
            worksheet.write('L25', adjPenalty, workbookFormat(workbook, defaultFormat))
        elif (x == 'K'):
            worksheet.write('K25', penPaid, workbookFormat(workbook, defaultFormat))
        elif (x == 'J'):
            worksheet.write('J25', penTotal, workbookFormat(workbook, defaultFormat))
        else:
            worksheet.write('{}25'.format(x), dfCustomerLedger['accountSummary'][y],
                            workbookFormat(workbook, defaultFormat))

    worksheet.write('J26', dfCustomerLedger['accountSummary']['unappliedBalance'], workbookFormat(workbook, defaultFormat))

    worksheet.merge_range('H18:I18', 'GRAND TOTAL', workbookFormat(workbook, ledgerHeader))
    worksheet.merge_range('H20:I20', 'TOTAL PNV', workbookFormat(workbook, ledgerHeader))
    worksheet.merge_range('H21:I21', 'RFC', workbookFormat(workbook, stringFormat))
    worksheet.merge_range('H22:I22', 'PRINCIPAL', workbookFormat(workbook, defaultFormat))
    worksheet.merge_range('H23:I23', 'INTEREST', workbookFormat(workbook, defaultFormat))
    worksheet.merge_range('H25:I25', 'PENALTY', workbookFormat(workbook, stringFormat))
    worksheet.merge_range('H26:I26', 'ADVANCES', workbookFormat(workbook, stringFormat))

    worksheet.merge_range('A22:F22', 'OUTSTANDING BALANCE:', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(cnumbers(7, 23), headersOB):
        worksheet.merge_range('A{}:D{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    worksheet.merge_range('E23:F23', '=N21', workbookFormat(workbook, numFormat))
    worksheet.merge_range('E24:F24', '=N25', workbookFormat(workbook, numFormat))
    worksheet.merge_range('E26:F26', '=SUM(E23,E24)', workbookFormat(workbook, numFormat))
    worksheet.merge_range('E28:F28', '=J18', workbookFormat(workbook, numFormat))
    worksheet.merge_range('E29:F29', dfCustomerLedger['acctStat']['lastPaymentDate'], workbookFormat(workbook, numFormat))

    # for x, y in zip(numbers(6, 23), dataOB):
    #     if(x == 25):
    #         worksheet.write('E25', dfOBDetails[y], workbookFormat(workbook, defaultUnderlineFormat))
    #     elif(x == 26):
    #         worksheet.write('E26', dfOBDetails[y], workbookFormat(workbook, defaultUnderlineFormat))
    #     elif(y == 'totalPayment'):
    #         worksheet.merge_range('D28:E28'.format(x, x), dfOBDetails['totalPayment'], workbookFormat(workbook, defaultFormat))
    #     elif(y == 'lastPaymentDate'):
    #         worksheet.merge_range('D30:E30'.format(x, x), dfOBDetails[y], workbookFormat(workbook, defaultFormat))
    #     else:
    #         worksheet.merge_range('D{}:E{}'.format(x, x), dfOBDetails['lastPaymentDate'], workbookFormat(workbook, defaultFormat))

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}32:{}33'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    # worksheet.merge_range('A{}:R{}'.format(count + 1, count + 1), nodisplay, workbookFormat(workbook, entriesStyle))
    # worksheet.merge_range('A{}:C{}'.format(count + 2, count + 2), 'PAID', workbookFormat(workbook, docNameStyle))
    # worksheet.write('I{}'.format(count + 2), totalPrincipalPaid, workbookFormat(workbook, footerStyle))
    # worksheet.write('J{}'.format(count + 2), totalInterestPaid, workbookFormat(workbook, footerStyle))
    # worksheet.write('K{}'.format(count + 2), totalPenaltyPaid, workbookFormat(workbook, footerStyle))
    #
    # worksheet.merge_range('A{}:C{}'.format(count + 3, count + 3), 'BILLED', workbookFormat(workbook, docNameStyle))
    # worksheet.write('H{}'.format(count + 3), totalPenalty, workbookFormat(workbook, footerStyle))
    # worksheet.write('I{}'.format(count + 3), totalPrincipalBilled, workbookFormat(workbook, footerStyle))
    # worksheet.write('J{}'.format(count + 3), totalInterestBilled, workbookFormat(workbook, footerStyle))

    # for c in range(ord('G'), ord('N') + 1):
    #     worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}34:{}{})".format(chr(c), chr(c), count - 1),
    #                     merge_format4)
    writer.close()

    output.seek(0)
    print('sending spreadsheet')
    filename = "Customer Ledger {}.xlsx".format(loanId)
    return send_file(output, attachment_filename=filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=port)