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

app = Flask(__name__)
excel.init_excel(app)
# port = 5001
port = int(os.getenv("PORT"))

fmtDate = "%m/%d/%y"
fmtTime = "%I:%M %p"
now_utc = datetime.now(timezone('UTC'))
now_pacific = now_utc.astimezone(timezone('Asia/Manila'))
dateNow = now_pacific.strftime(fmtDate)
timeNow = now_pacific.strftime(fmtTime)

comNameStyle = {'font':'Gill Sans MT', 'font_size': '16','bold': True, 'align': 'left'}
docNameStyle = {'font':'Segeo UI', 'font_size': '8', 'bold': True, 'align': 'left'}
periodStyle = {'font':'Segeo UI', 'font_size': '8', 'align': 'left'}
ledgerDataStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'right'}
ledgerNameStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'left'}
undStyle = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'bold': True, 'underline': True}
generatedStyle = {'font':'Segeo UI', 'font_size': '8', 'align': 'right'}
headerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True}
textWrapHeader = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True}
entriesStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'bottom': 2, 'align': 'center'}
borderFormatStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'bottom': 2, 'align': 'center', 'num_format': '₱#,##0.00'}
topBorderStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'top': 2, 'align': 'center'}
footerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'right', 'num_format': '₱#,##0.00'}
sumStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'right'}
centerStyle = {'font':'Segeo UI', 'font_size': '7', 'bold': True, 'align': 'center'}
numFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'num_format': '#,##0.00'}
stringFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'left', 'num_format': '#,##0.00'}
defaultFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right'}
defaultUnderlineFormat = {'font':'Segeo UI', 'font_size': '7', 'align': 'right', 'bottom': 2}
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
    worksheet.merge_range('B{}:B{}'.format(counts + 4, counts + 5), 'DATE', merge_format7)
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

    url = 'https://api360.zennerslab.com/Service1.svc/collection'
    # url = 'https://rfc360-test.zennerslab.com/Service1.svc/collection'
    # url = 'http://localhost:15021/Service1.svc/collection'
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
        if (x == 'P'):
            worksheet.write('P7', 'HF', workbookFormat(workbook, headerStyle))
        elif (x == 'Q'):
            worksheet.write('Q7', 'DST', workbookFormat(workbook, headerStyle))
        elif (x == 'R'):
            worksheet.write('R7', 'NOTARIAL', workbookFormat(workbook, headerStyle))
        elif (x == 'S'):
            worksheet.write('S7', 'GCLI', workbookFormat(workbook, headerStyle))
        else:
            worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('P6:S6', 'UPFRONT CHARGES', workbookFormat(workbook, headerStyle))

    worksheet.merge_range('A{}:V{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    for c in range(ord('E'), ord('V') + 1):
        if (chr(c) == 'F'):
            worksheet.write('F{}'.format(count + 1), "=SUM(F8:F{})".format(count - 1), workbookFormat(workbook, sumStyle))
        elif (chr(c) == 'J'):
            worksheet.write('J{}'.format(count + 1), "=SUM(J8:J{})".format(count - 1), workbookFormat(workbook, sumStyle))
        elif (chr(c) == 'N'):
            worksheet.write('N{}'.format(count + 1), "=SUM(N8:N{})".format(count - 1), workbookFormat(workbook, sumStyle))
        elif (chr(c) == 'O'):
            worksheet.write('N{}'.format(count + 1), "=SUM(O8:O{})".format(count - 1), workbookFormat(workbook, sumStyle))
        elif (chr(c) == 'U'):
            worksheet.write('U{}'.format(count + 1), "")
        else:
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                            workbookFormat(workbook, footerStyle))

    writer.close()
    output.seek(0)

    print('sending spreadsheet')

    filename = "Collection Report {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)

# 2 sheets
@app.route("/agingReport", methods=['GET'])
# def newAgingReport():
#
#     output = BytesIO()
#
#     name = request.args.get('name')
#     date = request.args.get('date')
#
#     payload = {'date': date}
#
#     # url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport" #lambda-live
#     # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport"  # lambda-test
#     url = "http://localhost:6999/reports/accountingAgingReport" #lambda-localhost
#     # url ="https://report-cache.cfapps.io/accountingAging"
#
#     r = requests.post(url, json=payload)
#     data = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME",
#                "COLLECTOR", "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "CURR. TODAY",
#                "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "OVER 360"]
#
#     agingp1headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME", "COLLECTOR",
#                       "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "CURR. TODAY"]
#     agingp11headers = ["1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "OVER 360"]
#     agingp2headers = ["#", "PRINCPAL", "INTEREST", "PENALTY", "TOTAL", "ADV"]
#
#     agingp1DF = pd.DataFrame(data)
#     agingp2DF = pd.DataFrame(data).copy()
#
#     agingp1list1 = [len(i) for i in headers]
#     agingp2list1 = [len(i) for i in agingp2headers]
#
#     if agingp1DF.empty:
#         count1 = agingp1DF.shape[0] + 8
#         agingp1nodisplay = 'No Data'
#         agingp1DF = pd.DataFrame(pd.np.empty((0, 19)))
#         agingp1list2 = agingp1list1
#     else:
#         count1 = agingp1DF.shape[0] + 8
#         agingp1nodisplay = ''
#         agingp1DF['num'] = numbers(agingp1DF.shape[0])
#         astype(agingp1DF, 'term', int)
#         astype(agingp1DF, 'expiredTerm', int)
#         astype(agingp1DF, 'appId', int)
#         astype(agingp1DF, 'runningPNV', float)
#         astype(agingp1DF, 'runningMLV', float)
#         astype(agingp1DF, 'monthlyInstallment', float)
#         dfDateFormat(agingp1DF, 'fdd')
#         dfDateFormat(agingp1DF, 'lastPaymentDate')
#         agingp1DF['loanAccountNumber'] = agingp1DF['loanAccountNumber'].map(lambda x: x.lstrip("'"))
#         agingp1DF['lastPaymentDate'] = agingp1DF.lastPaymentDate.apply(lambda x: x.split(" ")[0])
#         agingp1DF = round(agingp1DF, 2)
#         agingp1DF = agingp1DF[["num", "channelName", "partnerCode", "outletCode", "appId", "loanAccountNumber", "fullName",
#                                "alias", "fdd", "lastPaymentDate", "term", "expiredTerm", "monthlyInstallment", "stats", "runningPNV", "runningMLV", "today",
#                                "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "total"]]
#         agingp1list2 = [max([len(str(s)) for s in agingp1DF[col].values]) for col in agingp1DF.columns]
#
#     if agingp2DF.empty:
#         count = agingp2DF.shape[0] + 8
#         agingp2nodisplay = 'No Data'
#         agingp2DF = pd.DataFrame(pd.np.empty((0, 19)))
#         agingp2list2 = agingp1list1
#     else:
#         count = agingp2DF.shape[0] + 8
#         agingp2nodisplay = ''
#         agingp2DF['num'] = numbers(agingp2DF.shape[0])
#         astype(agingp2DF, 'duePrincipal', float)
#         astype(agingp2DF, 'dueInterest', float)
#         astype(agingp2DF, 'duePenalty', float)
#         agingp2DF['loanAccountNumber'] = agingp2DF['loanAccountNumber'].map(lambda x: x.lstrip("'"))
#         agingp2DF = round(agingp2DF, 2)
#         agingp2DF['adv'] = '-'
#         agingp2DF = agingp2DF[["num", "duePrincipal", "dueInterest", "duePenalty", "total", "adv"]]
#         agingp2list2 = [max([len(str(s)) for s in agingp2DF[col].values]) for col in agingp2DF.columns]
#
#     # agingp1DF = agingp1DF.style.set_properties(**styles)
#     # agingp2DF = agingp2DF.style.set_properties(**styles)
#     agingp1DF.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="AgingP1", header=None)
#     agingp2DF.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="AgingP2", header=None)
#
#     workbook = writer.book
#
#     worksheetAgingP1 = writer.sheets["AgingP1"]
#
#     dataframeStyle(worksheetAgingP1, 'A', 'A', 8, count, workbookFormat(workbook, defaultFormat))
#     dataframeStyle(worksheetAgingP1, 'B', 'D', 8, count, workbookFormat(workbook, stringFormat))
#     dataframeStyle(worksheetAgingP1, 'E', 'E', 8, count, workbookFormat(workbook, defaultFormat))
#     dataframeStyle(worksheetAgingP1, 'F', 'H', 8, count, workbookFormat(workbook, stringFormat))
#     dataframeStyle(worksheetAgingP1, 'I', 'J', 8, count, workbookFormat(workbook, numFormat))
#     dataframeStyle(worksheetAgingP1, 'K', 'L', 8, count, workbookFormat(workbook, defaultFormat))
#     dataframeStyle(worksheetAgingP1, 'M', 'M', 8, count, workbookFormat(workbook, numFormat))
#     dataframeStyle(worksheetAgingP1, 'N', 'N', 8, count, workbookFormat(workbook, stringFormat))
#     dataframeStyle(worksheetAgingP1, 'O', 'Z', 8, count, workbookFormat(workbook, numFormat))
#
#     for col_num, value in enumerate(columnWidth(agingp1list1, agingp1list2)):
#         worksheetAgingP1.set_column(col_num, col_num, value)
#
#     worksheetAgingP1.freeze_panes(7, 0)
#
#     def alphabetRange(firstRange, secondRange):
#         alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
#         return alphaList
#
#     range1 = 'W'
#     range2 = 'X'
#     range3 = 'Z'
#     companyName = 'RFSC'
#     reportTitle = 'AGING REPORT'
#     branchName = 'Nationwide'
#     xldate_header = "As of {}".format(startDateFormat(date))
#
#     workSheet(workbook, worksheetAgingP1, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)
#
#     headersList = [i for i in agingp1headers]
#     headersList1 = [i for i in agingp11headers]
#
#     for x, y in zip(alphabetRange('A', 'Q'), headersList):
#         worksheetAgingP1.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))
#
#     worksheetAgingP1.merge_range('R6:Y6', 'PAST DUE', workbookFormat(workbook, headerStyle))
#
#     for x, y in zip(alphabetRange('R', 'Y'), headersList1):
#         worksheetAgingP1.write('{}7'.format(x), '{}'.format(y), workbookFormat(workbook, headerStyle))
#
#     worksheetAgingP1.merge_range('Z6:Z7', 'TOTAL DUE', workbookFormat(workbook, headerStyle))
#
#     worksheetAgingP1.merge_range('A{}:Z{}'.format(count1, count1), agingp1nodisplay, workbookFormat(workbook, entriesStyle))
#     worksheetAgingP1.merge_range('A{}:C{}'.format(count1 + 1, count1 + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))
#
#     worksheetAgingP1.write('M{}'.format(count1 + 1), "=SUM(M8:M{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
#     for c in range(ord('O'), ord('Z') + 1):
#         worksheetAgingP1.write('{}{}'.format(chr(c), count1 + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count1 - 1),
#                             workbookFormat(workbook, footerStyle))
#
#     worksheetAgingP2 = writer.sheets["AgingP2"]
#
#     # workSheet(workbook, worksheetAgingP2, range1, range2, range3, xldate_header, name, companyName, reportTitle,
#     #           branchName)
#
#     dataframeStyle(worksheetAgingP2, 'A', 'A', 8, count, workbookFormat(workbook, defaultFormat))
#     dataframeStyle(worksheetAgingP2, 'B', 'F', 8, count, workbookFormat(workbook, numFormat))
#
#
#     for col_num, value in enumerate(columnWidth(agingp2list1, agingp2list2)):
#         worksheetAgingP2.set_column(col_num, col_num, value)
#
#     # headersList2 = [i for i in agingp2headers]
#
#     worksheetAgingP2.merge_range('A5:B5', 'AGING REPORT', workbookFormat(workbook, docNameStyle))
#     worksheetAgingP2.merge_range('C5:I5', 'NATIONWIDE', workbookFormat(workbook, centerStyle))
#     worksheetAgingP2.merge_range('J5:K5', 'PAGE 2 OF 2', workbookFormat(workbook, docNameStyle))
#
#     worksheetAgingP2.freeze_panes(7, 0)
#
#     worksheetAgingP2.merge_range('A6:A7', '#', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.merge_range('B6:D6', 'PAST DUE BREAKDOWN', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.write('B7', 'PRINCIPAL', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.write('C7', 'INTEREST', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.write('D7', 'PENALTY', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.merge_range('E6:E7', 'TOTAL', workbookFormat(workbook, headerStyle))
#     worksheetAgingP2.merge_range('F6:F7', 'ADV', workbookFormat(workbook, headerStyle))
#
#     worksheetAgingP2.merge_range('A{}:F{}'.format(count, count), agingp2nodisplay, workbookFormat(workbook, entriesStyle))
#     # worksheetAgingP2.write('A{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)
#
#     for c in range(ord('B'), ord('E') + 1):
#         worksheetAgingP2.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
#                                workbookFormat(workbook, footerStyle))
#     # the writer has done its job
#     writer.close()
#
#     # go back to the beginning of the stream
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Aging Report as of {}.xlsx".format(date)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newAgingReport", methods=['GET'])
def newAgingReport2():

    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')

    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport" #lambda-live
    # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport"  # lambda-test
    # url = "http://localhost:6999/reports/accountingAgingReport" #lambda-localhost
    # url ="https://report-cache.cfapps.io/accountingAging"

    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME",
               "COLLECTOR", "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "CURR. TODAY",
               "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "OVER 360"]

    agingp1headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME", "COLLECTOR",
                      "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "CURR. TODAY"]
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
        dfDateFormat(agingp1DF, 'fdd')
        dfDateFormat(agingp1DF, 'lastPaymentDate')
        agingp1DF['loanAccountNumber'] = agingp1DF['loanAccountNumber'].map(lambda x: x.lstrip("'"))
        agingp1DF['lastPaymentDate'] = agingp1DF.lastPaymentDate.apply(lambda x: x.split(" ")[0])
        # agingp1DF['adv'] = '-'
        agingp1DF = round(agingp1DF, 2)
        agingp1DF = agingp1DF[["num", "channelName", "partnerCode", "outletCode", "appId", "loanAccountNumber", "fullName",
                               "alias", "fdd", "lastPaymentDate", "term", "expiredTerm", "monthlyInstallment", "stats", "runningPNV", "runningMLV", "today",
                               "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty"]]
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
    dataframeStyle(worksheetAgingP1, 'O', 'Z', 8, count1, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AA8:AA{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AB8:AB{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    worksheetAgingP1.set_column('AC8:AC{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    # worksheetAgingP1.set_column('AD8:AD{}'.format(count1 - 1), None, workbookFormat(workbook, numFormat))
    for col_num, value in enumerate(columnWidth(agingp1list1, agingp1list2)):
        worksheetAgingP1.set_column(col_num, col_num, value)

    worksheetAgingP1.freeze_panes(7, 0)

    def alphabetRange(firstRange, secondRange):
        alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
        return alphaList

    range1 = 'Z'
    range2 = 'AA'
    range3 = 'AC'
    companyName = 'RFSC'
    reportTitle = 'AGING REPORT'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheetAgingP1, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in agingp1headers]
    headersList1 = [i for i in agingp11headers]

    for x, y in zip(alphabetRange('A', 'Q'), headersList):
        worksheetAgingP1.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('R6:Y6', 'PAST DUE', workbookFormat(workbook, headerStyle))

    for x, y in zip(alphabetRange('R', 'Y'), headersList1):
        worksheetAgingP1.write('{}7'.format(x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('Z6:Z7', 'TOTAL DUE', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.merge_range('AA6:AC6', 'PAST DUE BREAKDOWN', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AA7', 'PRINCIPAL', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AB7', 'INTEREST', workbookFormat(workbook, headerStyle))
    worksheetAgingP1.write('AC7', 'PENALTY', workbookFormat(workbook, headerStyle))
    # worksheetAgingP1.merge_range('AD6:AD7', 'ADV', workbookFormat(workbook, headerStyle))

    worksheetAgingP1.merge_range('A{}:AC{}'.format(count1, count1), agingp1nodisplay, workbookFormat(workbook, entriesStyle))
    worksheetAgingP1.merge_range('A{}:C{}'.format(count1 + 1, count1 + 1), 'GRAND TOTAL:', workbookFormat(workbook, docNameStyle))

    worksheetAgingP1.write('M{}'.format(count1 + 1), "=SUM(M8:M{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    for c in range(ord('O'), ord('Z') + 1):
        worksheetAgingP1.write('{}{}'.format(chr(c), count1 + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count1 - 1),
                            workbookFormat(workbook, footerStyle))

    worksheetAgingP1.write('AA{}'.format(count1 + 1), "=SUM(AA8:AA{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AB{}'.format(count1 + 1), "=SUM(AB8:AB{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    worksheetAgingP1.write('AC{}'.format(count1 + 1), "=SUM(AC8:AC{})".format(count1 - 1), workbookFormat(workbook, footerStyle))
    # worksheetAgingP1.write('AD{}'.format(count1 + 1), "=SUM(AD8:AD{})".format(count1 - 1), workbookFormat(workbook, footerStyle))


    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/accountingAgingReport", methods=['GET'])
def accountingAgingReport():

    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')

    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport" #lambda-live
    # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/accountingAgingReport"  # lambda-test
    # url = "http://localhost:6999/reports/accountingAgingReport" #lambda-localhost
    # url ="https://report-cache.cfapps.io/accountingAging"

    r = requests.get(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "COLLECTOR", "CLIENT'S NAME", "MOBILE #", "ADDRESS", "LOAN ACCT. #", "TODAY", "1-30",
               "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & OVER", "TOTAL", "MATURED",
               "DUE PRINCIPAL", "DUE INTEREST", "DUE PENALTY"]
    df = pd.DataFrame(data)
    list1 = [len(i) for i in headers]

    if df.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 19)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        df['loanAccountNumber'] = df['loanAccountNumber'].map(lambda x: x.lstrip("'"))
        df = round(df, 2)
        df['num'] = numbers(df.shape[0])
        df = df[["num", "collector", "fullName", "mobile", "address", "loanAccountNumber", "today","1-30", "31-60", "61-90",
                 "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "matured", "principal",
                 "interest", "penalty"]]
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

    range1 = 'Q'
    range2 = 'R'
    range3 = 'T'
    companyName = 'RFSC'
    reportTitle = 'Accounting Aging Report'
    branchName = 'Nationwide'
    xldate_header = "As of {}".format(startDateFormat(date))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), merge_format7)

    worksheet.merge_range('A{}:T{}'.format(count, count), nodisplay, merge_format6)
    worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)

    for c in range(ord('G'), ord('T') + 1):
        if (chr(c) == 'Q'):
            worksheet.write('Q{}'.format(count + 1), "=SUM(Q8:Q{})".format(count - 1), merge_format8)
        else:
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                                merge_format4)

    # the writer has done its job
    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "Aging Report (Accounting) as of {}.xlsx".format(date)
    return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/operationAgingReport", methods=['GET'])
def operationAgingReport():

    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')

    payload = {'date': date}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport" #lambda-live
    # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport" #lambda-test
    # url = "http://localhost:6999/reports/operationAgingReport" #lambda-localhost
    # url = "https://report-cache.cfapps.io/operationAging"
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

@app.route("/newoperationAgingReport", methods=['GET'])
# def newoperationAgingReport():
#
#     output = BytesIO()
#
#     name = request.args.get('name')
#     date = request.args.get('date')
#
#     payload = {'date': date}
#
#     # url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport" #lambda-live
#     url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/operationAgingReport"  # lambda-test
#     # url = "http://localhost:6999/reports/operationAgingReport" #lambda-localhost
#     r = requests.post(url, json=payload)
#     data = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     df = pd.DataFrame(data['operationAgingReportJson'])
#     df['appId'] = df['appId'].astype(int)
#     df.sort_values(by=['appId'])
#
#     if df.empty:
#         count = df.shape[0] + 9
#         nodisplay = 'No Data'
#         totalsum = 0
#         principalsum = 0
#         interestsum = 0
#         penaltysum = 0
#         df = pd.DataFrame(pd.np.empty((0, 28)))
#     else:
#         count = df.shape[0] + 9
#         nodisplay = ''
#         totalsum = pd.Series(df['total']).sum()
#         principalsum = pd.Series(df['duePrincipal']).sum()
#         interestsum = pd.Series(df['dueInterest']).sum()
#         penaltysum = pd.Series(df['duePenalty']).sum()
#         df['loanaccountNumber'] = df['loanaccountNumber'].map(lambda x: x.lstrip("'"))
#         df['fdd'] = pd.to_datetime(df['fdd'])
#         df['fdd'] = df['fdd'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
#         df = df[["appId", "loanaccountNumber", "fullName", "mobile", "address", "term", "fdd", "status", "PNV",
#                  "MLV", "bPNV", "bMLV", "mi", "notDue", "matured", "today", "1-30", "31-60", "61-90", "91-120",
#                  "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal", "dueInterest", "duePenalty"]]
#
#     df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)
#
#     workbook = writer.book
#     merge_format1 = workbook.add_format({'align': 'center'})
#     merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
#     merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
#     merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
#     merge_format5 = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True})
#     xldate_header = "As of {}".format(date)
#
#     worksheet = writer.sheets["Sheet_1"]
#
#     list1 = [len(i) for i in df.columns.values]
#     # list1 = np.array(headerlen)
#
#     if df.empty:
#         list2 = list1
#     else:
#         list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
#
#     def function(list1, list2):
#         list3 = [max(value) for value in zip(list1, list2)]
#         return list3
#
#     for col_num, value in enumerate(function(list1, list2)):
#         worksheet.set_column(col_num, col_num, value + 1)
#
#     worksheet.merge_range('A1:W1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheet.merge_range('A2:W2', 'RFC360 Kwikredit', merge_format1)
#     worksheet.merge_range('A3:W3', 'Aging Report (Operations)', merge_format3)
#     worksheet.merge_range('A4:W4', xldate_header, merge_format1)
#
#     worksheet.merge_range('A6:A7', 'Loan', merge_format5)
#     worksheet.merge_range('B6:B7', 'Product Type', merge_format5)
#     worksheet.merge_range('C6:C7', 'Customer Name', merge_format5)
#     worksheet.merge_range('D6:D7', 'Address', merge_format5)
#     worksheet.merge_range('E6:E7', 'CCI Officer', merge_format5)
#     worksheet.merge_range('F6:F7', 'FDD', merge_format5)
#     worksheet.merge_range('G6:G7', 'Term', merge_format5)
#     worksheet.merge_range('H6:H7', 'Exp Term', merge_format5)
#     worksheet.merge_range('I6:I7', 'MI', merge_format5)
#     worksheet.merge_range('J6:J7', 'Status', merge_format5)
#     worksheet.merge_range('K6:K7', 'Restructed', merge_format5)
#     worksheet.merge_range('L6:L7', 'OB', merge_format5)
#     worksheet.merge_range('M6:M7', 'Not Due', merge_format5)
#     worksheet.merge_range('N6:N7', 'Current Today', merge_format5)
#     worksheet.merge_range('O6:V6', 'PAST DUE', merge_format5)
#     worksheet.write('O7', '1-30', merge_format5)
#     worksheet.write('P7', '31-60', merge_format5)
#     worksheet.write('Q7', '61-90', merge_format5)
#     worksheet.write('R7', '91-120', merge_format5)
#     worksheet.write('S7', '121-150', merge_format5)
#     worksheet.write('T7', '151-180', merge_format5)
#     worksheet.write('U7', '181-360', merge_format5)
#     worksheet.write('V7', 'OVER 360', merge_format5)
#     worksheet.merge_range('W6:W7', 'Total Due', merge_format5)
#
#     worksheet.merge_range('A{}:W{}'.format(count - 1, count - 1), nodisplay, merge_format1)
#     worksheet.merge_range('W{}:X{}'.format(count + 1, count + 1), 'TOTAL', merge_format3)
#     worksheet.write('Y{}'.format(count + 1), totalsum, merge_format4)
#     worksheet.write('Z{}'.format(count + 1), principalsum, merge_format4)
#     worksheet.write('AA{}'.format(count + 1), interestsum, merge_format4)
#     worksheet.write('AB{}'.format(count + 1), penaltysum, merge_format4)
#     worksheet.merge_range('A{}:W{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
#     worksheet.merge_range('A{}:W{}'.format(count + 4, count + 5), name, merge_format2)
#     worksheet.merge_range('A{}:W{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
#                           merge_format2)
#
#     # the writer has done its job
#     writer.close()
#
#     # go back to the beginning of the stream
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Aging Report (Operations) as of {}.xlsx".format(date)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmemoreport2", methods=['GET'])
# def newmemoreport2():
#
#     output = BytesIO()
#
#     name = request.args.get('name')
#     dateStart = request.args.get('startDate')
#     dateEnd = request.args.get('endDate')
#
#     payload = {'startDate': dateStart, 'endDate': dateEnd}
#
#     # url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport" #lambda-live
#     url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport"  # lambda-test
#     # url = "http://localhost:6999/reports/memoreport" #lambda-localhost
#
#     r = requests.post(url, json=payload)
#     data = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["App ID", "Loan Account Number", "Customer Name", "Sub Product", "Memo Type", "Purpose", "Amount",
#                "Status", "Date", "Created By", "Approved By", "Approved Remarks"]
#
#     creditDf = pd.DataFrame(data['Credit'])
#     if creditDf.empty:
#         countCredit = creditDf.shape[0] + 8
#         nodisplayCredit = 'Nothing to display'
#         sumCredit = 0
#         creditDf = pd.DataFrame(pd.np.empty((0, 12)))
#     else:
#         countCredit = creditDf.shape[0] + 8
#         nodisplayCredit = ''
#         sumCredit = pd.Series(creditDf['amount']).sum()
#         creditDf.sort_values(by=['appId'], inplace=True)
#         creditDf['loanAccountNo'] = creditDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
#         creditDf['approvedDate'] = pd.to_datetime(creditDf['approvedDate'])
#         creditDf['approvedDate'] = creditDf['approvedDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
#         creditDf = creditDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
#                              "status", "date", "createdBy", "approvedBy", "approvedRemark"]]
#
#     debitDf = pd.DataFrame(data['Debit'])
#     if debitDf.empty:
#         countDebit = debitDf.shape[0] + 8
#         nodisplayDebit = 'Nothing to display'
#         sumDebit = 0
#         debitDf = pd.DataFrame(pd.np.empty((0, 12)))
#     else:
#         countDebit = debitDf.shape[0] + 8
#         nodisplayDebit = ''
#         sumDebit = pd.Series(debitDf['amount']).sum()
#         debitDf.sort_values(by=['appId'], inplace=True)
#         debitDf['loanAccountNo'] = debitDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
#         debitDf['approvedDate'] = pd.to_datetime(debitDf['approvedDate'])
#         debitDf['approvedDate'] = debitDf['approvedDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
#         debitDf = debitDf[["appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
#                            "status", "date", "createdBy", "approvedBy", "approvedRemark"]]
#
#
#     creditDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Credit", header=headers)
#     debitDf.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Debit", header=headers)
#
#     workbook = writer.book
#     merge_format1 = workbook.add_format({'align': 'center'})
#     merge_format2 = workbook.add_format({'bold': True, 'align': 'left'})
#     merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
#     merge_format4 = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'red', 'align': 'right'})
#     xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
#
#     worksheetCredit = writer.sheets["Credit"]
#
#     list1 = [len(i) for i in headers]
#     # list1 = np.array(headerlen)
#
#     if creditDf.empty:
#         list2 = list1
#     else:
#         list2 = [max([len(str(s)) for s in creditDf[col].values]) for col in creditDf.columns]
#
#     def function(list1, list2):
#         list3 = [max(value) for value in zip(list1, list2)]
#         return list3
#
#     for col_num, value in enumerate(function(list1, list2)):
#         worksheetCredit.set_column(col_num, col_num, value + 1)
#
#     worksheetCredit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheetCredit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
#     worksheetCredit.merge_range('A3:L3', 'Memo Report(Credit)', merge_format3)
#     worksheetCredit.merge_range('A4:L4', xldate_header, merge_format1)
#     worksheetCredit.merge_range('A{}:L{}'.format(countCredit - 1, countCredit - 1), nodisplayCredit, merge_format1)
#     worksheetCredit.merge_range('E{}:F{}'.format(countCredit + 1, countCredit + 1), 'TOTAL AMOUNT', merge_format3)
#     worksheetCredit.write('G{}'.format(countCredit + 1), sumCredit, merge_format4)
#     worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 3, countCredit + 3), 'Report Generated By :', merge_format2)
#     worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 4, countCredit + 5), name, merge_format2)
#     worksheetCredit.merge_range('A{}:L{}'.format(countCredit + 7, countCredit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
#                           merge_format2)
#
#     worksheetDebit = writer.sheets["Debit"]
#
#     list1 = [len(i) for i in headers]
#     # list1 = np.array(headerlen)
#
#     if debitDf.empty:
#         list2 = list1
#     else:
#         list2 = [max([len(str(s)) for s in debitDf[col].values]) for col in debitDf.columns]
#
#     def function(list1, list2):
#         list3 = [max(value) for value in zip(list1, list2)]
#         return list3
#
#     for col_num, value in enumerate(function(list1, list2)):
#         worksheetDebit.set_column(col_num, col_num, value + 1)
#
#     worksheetDebit.merge_range('A1:L1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheetDebit.merge_range('A2:L2', 'RFC360 Kwikredit', merge_format1)
#     worksheetDebit.merge_range('A3:L3', 'Memo Report(Debit)', merge_format3)
#     worksheetDebit.merge_range('A4:L4', xldate_header, merge_format1)
#     worksheetDebit.merge_range('A{}:L{}'.format(countDebit - 1, countDebit - 1), nodisplayDebit, merge_format1)
#     worksheetDebit.merge_range('E{}:F{}'.format(countDebit + 1, countDebit + 1), 'TOTAL AMOUNT', merge_format3)
#     worksheetDebit.write('G{}'.format(countDebit + 1), sumDebit, merge_format4)
#     worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 3, countDebit + 3), 'Report Generated By :', merge_format2)
#     worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 4, countDebit + 5), name, merge_format2)
#     worksheetDebit.merge_range('A{}:L{}'.format(countDebit + 7, countDebit + 7), 'Date & Time Report Generation ({})'.format(dateNow),
#                           merge_format2)
#
#     # the writer has done its job
#     writer.close()
#
#     # go back to the beginning of the stream
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmemoreport", methods=['GET'])
def newmemoreport():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport" #lambda-live
    # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/memoreport"  # lambda-test
    # url = "http://localhost:6999/reports/memoreport" #lambda-localhost

    r = requests.post(url, json=payload)
    data = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S", "SUB PRODUCT", "MEMO TYPE", "PURPOSE", "AMOUNT",
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
        creditDf.sort_values(by=['appId'], inplace=True)
        creditDf['loanAccountNo'] = creditDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        creditDf['date'] = creditDf.date.apply(lambda x: x.split(" ")[0])
        dfDateFormat(creditDf, 'approvedDate')
        dfDateFormat(creditDf, 'date')
        creditDf['num'] = numbers(creditDf.shape[0])
        creditDf = creditDf[["num", "appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]
        creditlist2 = [max([len(str(s)) for s in creditDf[col].values]) for col in creditDf.columns]

    debitDf = pd.DataFrame(data['Debit'])
    if debitDf.empty:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = 'No Data'
        debitDf = pd.DataFrame(pd.np.empty((0, 14)))
        debitlist2 = list1
    else:
        countDebit = debitDf.shape[0] + 8
        nodisplayDebit = ''
        debitDf.sort_values(by=['appId'], inplace=True)
        debitDf['loanAccountNo'] = debitDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        debitDf['date'] = creditDf.date.apply(lambda x: x.split(" ")[0])
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

@app.route("/memoreport", methods=['GET'])
# def memoreport():
#     output = BytesIO()
#
#     dateStart = request.args.get('startDate')
#     dateEnd = request.args.get('endDate')
#     payload = {'startDate': dateStart, 'endDate': dateEnd}
#
#     # url = 'https://api360.zennerslab.com/Service1.svc/getMemoReport'
#     url = 'https://rfc360-test.zennerslab.com/Service1.svc/getMemoReport'
#     r = requests.post(url, json=payload)
#     data = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["App ID", "Loan Account No", "Full Name", "Mobile Number", "Sub Product", "Memo Type", "Purpose", "Amount",
#                "Status", "Date Created", "Created By", "Remarks", "Approved Date", "Approved By", "Approved Remarks"]
#     df = pd.DataFrame(data['getMemoReportResult'])
#     df['loanId'] = df['loanId'].astype(int)
#     df.sort_values(by=['loanId'], inplace=True)
#     df['approvedDate'] = pd.to_datetime(df['approvedDate'])
#     df['approvedDate'] = df['approvedDate'].dt.strftime('%m/%d/%Y')
#
#     df = df[["loanId", "loanAccountNo", "fullName", "mobileNo", "subProduct", "memoType", "purpose", "amount",
#              "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]
#
#     df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)
#
#     workbook = writer.book
#     merge_format1 = workbook.add_format({'align': 'center'})
#     merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
#     xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
#
#     worksheet = writer.sheets["Sheet_1"]
#
#     list1 = [len(i) for i in headers]
#     # list1 = np.array(headerlen)
#
#     if df.empty:
#         list2 = list1
#     else:
#         list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
#
#     def function(list1, list2):
#         list3 = [max(value) for value in zip(list1, list2)]
#         return list3
#
#     for col_num, value in enumerate(function(list1, list2)):
#         worksheet.set_column(col_num, col_num, value + 1)
#
#     worksheet.merge_range('A1:O1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheet.merge_range('A2:O2', 'RFC360 Kwikredit', merge_format1)
#     worksheet.merge_range('A3:O3', 'Memo Report', merge_format3)
#     worksheet.merge_range('A4:O4', xldate_header, merge_format1)
#
#     # the writer has done its job
#     writer.close()
#
#     # go back to the beginning of the stream
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Memo Report {}-{}.xlsx".format(dateStart, dateEnd)
#     return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/tat", methods=['GET'])
def tat():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/newtat" #lambda-live
    # url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/newtat" #lambda-test
    # url = "http://localhost:6999/newtat" #lambda-localhost

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
    dataframeStyle(worksheetStandard, 'E', 'G', 8, countStandard, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetStandard, 'H', 'J', 8, countStandard, workbookFormat(workbook, stringFormat))
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

    dataframeStyle(worksheetStandard, 'A', 'B', 8, countStandard, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheetStandard, 'C', 'D', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetStandard, 'E', 'G', 8, countStandard, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheetStandard, 'H', 'J', 8, countStandard, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheetStandard, 'K', 'X', 8, countStandard, workbookFormat(workbook, defaultFormat))

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

@app.route("/oldtat", methods=['GET'])
# def oldtat():
#     output = BytesIO()
#
#     dateStart = request.args.get('startDate')
#     dateEnd = request.args.get('endDate')
#     payload = {'startDate': dateStart, 'endDate': dateEnd}
#
#     # url = "https://3l8yr5jb35.execute-api.us-east-1.amazonaws.com/latest/reports/newtat" #lambda-live
#     url = "https://rekzfwhmj8.execute-api.us-east-1.amazonaws.com/latest/reports/newtat"  # lambda-test
#     # url = "http://localhost:6999/newtat" #lambda-localhost
#
#     r = requests.post(url, json=payload)
#     data = r.json()
#     standard = data['standard']
#     returned = data['return']
#
#     standard_df = pd.read_csv(StringIO(standard))
#     returned_df = pd.read_csv(StringIO(returned))
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     standard_df.to_excel(writer, sheet_name="Standard", index=False)
#     returned_df.to_excel(writer, sheet_name="Returned", index=False)
#
#     writer.close()
#     output.seek(0)
#
#     filename = "TAT {}-{}.xlsx".format(dateStart, dateEnd)
#     return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/unappliedbalances", methods=['GET'])
def get_uabalances():
    output = BytesIO()

    name = request.args.get('name')
    date = request.args.get('date')
    payload = {}
    url = "https://api360.zennerslab.com/Service1.svc/accountDueReportJSON"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/accountDueReportJSON"
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
        df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['loanId'], inplace=True)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['dueDate'] = pd.to_datetime(df['dueDate'])
        df['dueDate'] = df['dueDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'dueDate')
        df = df[["num", "loanId", "loanAccountNo", "name", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]
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
# def get_data():
#
#     output = BytesIO()
#
#     name = request.args.get('name')
#     dateStart = request.args.get('startDate')
#     dateEnd = request.args.get('endDate')
#
#     payload = {'startDate': dateStart, 'endDate': dateEnd}
#     # url = "https://api360.zennerslab.com/Service1.svc/DCCRjson"
#     url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjson"
#     r = requests.post(url, json=payload)
#     data_json = r.json()
#     sortData = sorted(data_json['DCCRjsonResult'], key=lambda d: d['postedDate'], reverse=False)
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["LOAN ACCT. #", "CLIENT'S", "MOBILE NUMBER", "OR #", "OR DATE", "NET CASH",
#                "PAYMENT SOURCE"]
#     df = pd.DataFrame(sortData)
#     list1 = [len(i) for i in headers]
#
#     if df.empty:
#         count = df.shape[0] + 8
#         nodisplay = 'No Data'
#         df = pd.DataFrame(pd.np.empty((0, 7)))
#         list2 = list1
#     else:
#         count = df.shape[0] + 8
#         nodisplay = ''
#         df["customerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
#         df['amount'] = df['amount'].astype(float)
#         df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
#         df['postedDate'] = pd.to_datetime(df['postedDate'])
#         df['postedDate'] = df['postedDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
#         df = df[['loanAccountNo', 'customerName', 'mobileNo', 'orNo', "postedDate", "amount",
#                  "paymentSource"]]
#         list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
#
#     df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)
#
#     workbook = writer.book
#     merge_format2 = workbook.add_format(docNameStyle)
#     merge_format4 = workbook.add_format(footerStyle)
#     merge_format6 = workbook.add_format(entriesStyle)
#     merge_format7 = workbook.add_format(headerStyle)
#     xldate_header = "{} to {}".format(dateStart, dateEnd)
#
#     worksheet = writer.sheets["Sheet_1"]
#
#     for col_num, value in enumerate(columnWidth(list1, list2)):
#         worksheet.set_column(col_num, col_num, value + 1)
#
#     range1 = 'D'
#     range2 = 'E'
#     range3 = 'G'
#     companyName = 'Radiowealth Financial Services Corporation'
#     reportTitle = 'Daily Cash/Check Report'
#     branchName = 'Nationwide'
#     workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)
#
#     headersList = [i for i in headers]
#
#     for x, y in zip(alphabet(range3), headersList):
#         worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), merge_format7)
#
#     worksheet.merge_range('A{}:G{}'.format(count, count), nodisplay, merge_format6)
#     worksheet.merge_range('A{}:B{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)
#
#     worksheet.write('F{}'.format(count + 1), '=SUM(F8:F{})'.format(count - 1), merge_format4)
#
#     writer.close()
#
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newdccr", methods=['GET'])
def get_data1():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://api360.zennerslab.com/Service1.svc/DCCRjsonNew"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjsonNew"
    # url = "http://localhost:15021/Service1.svc/DCCRjsonNew"
    r = requests.post(url, json=payload)
    data_json = r.json()

    sortData = sorted(data_json['DCCRjsonNewResult'], key=lambda d: d['postedDate'], reverse=False)
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "COLLECTOR", "DATE", "OR #", "CHECK #", "DATE DEPOSITED", "AMT DEPOSITED", "PAYMENT TYPE",
               "LOAN ACCT. #", "CUSTOMER NAME", "TOTAL", "CASH", "CHECK", "PRINCIPAL", "INTEREST", "ADVANCES", "PENALTY",
               "GIBCO", "HF", "DST", "PF", "NOTARIALSS", "GCLI", "OTHERSS", "AMOUNT"]
    df = pd.DataFrame(sortData)
    df1 = pd.DataFrame(sortData).copy()
    list1 = [len(i) for i in headers]

    if df.empty or df1.empty:
        count = df.shape[0] + 8
        nodisplay = 'No Data'
        df = pd.DataFrame(pd.np.empty((0, 25)))
        dfCashcount = 0
        dfEcpaycount = 0
        dfBCcount = 0
        dfBankcount = 0
        dfCheckcount = 0
        df1['num1'] = ''
        dfCash = pd.DataFrame(pd.np.empty((0, 25)))
        dfEcpay = pd.DataFrame(pd.np.empty((0, 25)))
        dfBC = pd.DataFrame(pd.np.empty((0, 25)))
        dfBank = pd.DataFrame(pd.np.empty((0, 25)))
        dfCheck = pd.DataFrame(pd.np.empty((0, 25)))
        df2 = pd.DataFrame(pd.np.empty((0, 25)))
        list2 = list1
    else:
        count = df.shape[0] + 8
        nodisplay = ''
        conditions = [(df['paymentSource'] == 'Check')]
        dfDateFormat(df, 'orDate')
        dfDateFormat(df, 'paymentDate')
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['total'] = np.select(conditions, [df['paymentCheck']], default=df['amount'])
        df['total1'] = np.select(conditions, [df['paymentCheck']], default=df['amount'])
        df1['total'] = np.select(conditions, [df1['paymentCheck']], default=df1['amount'])
        diff = df['total'] - (df['paidPrincipal'] + df['paidInterest'] + df['paidPenalty'])
        df['advances'] = round(diff, 2)
        df['gibco'] = 0
        df['hf'] = 0
        df['dst'] = 0
        df['pf'] = 0
        df['notarial'] = 0
        df['gcli'] = 0
        df['otherFees'] = 0
        df['amount1'] = 0
        df['description'] = ''
        df['num'] = numbers(df.shape[0])
        df1['num1'] = ''
        df['num1'] = ''
        df = round(df, 2)
        df1 = round(df, 2)
        df1 = df1.sort_values(by=['paymentSource'])
        dfCash = df1.loc[df1['paymentSource'] == 'Cash'].copy()
        dfEcpay = df1.loc[df1['paymentSource'] == 'Ecpay'].copy()
        dfBC = df1.loc[df1['paymentSource'] == 'Bayad Center'].copy()
        dfCheck = df1.loc[df1['paymentSource'] == 'Check'].copy()
        dfBank = df1.loc[df1['paymentSource'].isin(['Landbank','PNB','BDO','Metrobank','Unionbank'])].copy()

        dfCashcount = dfCash.shape[0]
        dfEcpaycount = dfEcpay.shape[0]
        dfBCcount = dfBC.shape[0]
        dfBankcount = dfBank.shape[0]
        dfCheckcount = dfCheck.shape[0]

        dfCash['dfCashnum'] = numbers(dfCashcount)
        dfEcpay['dfEcpaynum'] = numbers(dfEcpaycount)
        dfBC['dfBCnum'] = numbers(dfBCcount)
        dfBank['dfBanknum'] = numbers(dfBankcount)
        dfCheck['dfChecknum'] = numbers(dfCheckcount)
        df = df[['num', 'collector', 'orDate', 'orNo', 'checkNo', 'paymentDate', 'total1', 'paymentSource',
                 'loanAccountNo', 'customerName', 'total', 'amount', 'paymentCheck', 'paidPrincipal', 'paidInterest',
                 'advances', 'paidPenalty', 'gibco', 'hf', 'dst', 'pf', 'notarial', 'gcli', 'otherFees', 'amount1']]
        dfCash = dfCash[['dfCashnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfEcpay = dfEcpay[['dfEcpaynum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBC = dfBC[['dfBCnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBank = dfBank[['dfBanknum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfCheck = dfCheck[['dfChecknum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
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
    elif (dfEcpaycount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfBCcount + dfBankcount + 25)
    elif (dfBCcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBankcount + 25)
    elif (dfBankcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 25)
    elif (dfCheckcount <= 0):
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
    else:
        dfwriter(dfCash.to_excel, writer, count + 5)
        dfwriter(dfEcpay.to_excel, writer, count + dfCashcount + 10)
        dfwriter(dfBC.to_excel, writer, count + dfCashcount + dfEcpaycount + 15)
        dfwriter(dfBank.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + 20)
        dfwriter(dfCheck.to_excel, writer, count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 25)

    dfwriter(df2.to_excel, writer, count + count + 2)

    dataframeStyle(worksheet, 'A', 'A', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'B', 'B', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'C', 'C', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'D', 'E', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'F', 'G', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'H', 'J', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'K', 'Y', 8, count, workbookFormat(workbook, numFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        if(col_num == 3):
            worksheet.set_column(3, 3, 12)
        else:
            worksheet.set_column(col_num, col_num, value)

    # worksheet.freeze_panes(5, 0)

    range1 = 'V'
    range2 = 'W'
    range3 = 'Y'
    companyName = 'RFSC'
    reportTitle = 'DAILY CASH/CHECK COLLECTION (NET)'
    branchName = 'Nationwide'
    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    headersList = [i for i in headers]

    for x, y in zip(alphabet('Y'), headersList):
        if (x == 'K'):
            worksheet.write('K7', 'TOTAL', workbookFormat(workbook, headerStyle))
        elif (x == 'L'):
            worksheet.write('L7', 'CASH', workbookFormat(workbook, headerStyle))
        elif (x == 'M'):
            worksheet.write('M7', 'CHECK', workbookFormat(workbook, headerStyle))
        elif (x == 'N'):
            worksheet.write('N7', 'PRINCIPAL', workbookFormat(workbook, headerStyle))
        elif (x == 'O'):
            worksheet.write('O7', 'INTEREST', workbookFormat(workbook, headerStyle))
        elif (x == 'P'):
            worksheet.write('P7', 'ADVANCES', workbookFormat(workbook, headerStyle))
        elif (x == 'Q'):
            worksheet.write('Q7', 'PENALTY\n(5%)', workbookFormat(workbook, textWrapHeader))
        elif (x == 'R'):
            worksheet.write('R7', 'GIBCO', workbookFormat(workbook, headerStyle))
        elif (x == 'S'):
            worksheet.write('S7', 'HF', workbookFormat(workbook, headerStyle))
        elif (x == 'T'):
            worksheet.write('T7', 'DST', workbookFormat(workbook, headerStyle))
        elif (x == 'U'):
            worksheet.write('U7', 'PF', workbookFormat(workbook, headerStyle))
        elif (x == 'V'):
            worksheet.write('V7', 'NOTARIAL\nFEE', workbookFormat(workbook, textWrapHeader))
        elif (x == 'W'):
            worksheet.write('W7', 'GCLI', workbookFormat(workbook, headerStyle))
        elif (x == 'X'):
            worksheet.merge_range('X6:X7'.format(x, x), 'OTHER\nFEES', workbookFormat(workbook, textWrapHeader))
        else:
            worksheet.merge_range('{}6:{}7'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    worksheet.merge_range('K6:M6', 'AMOUNT', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('N6:Q6', 'LOAN REPAYMENT', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('R6:W6', 'ONE TIME PAYMENT', workbookFormat(workbook, headerStyle))

    worksheet.merge_range('J{}:Y{}'.format(count, count), nodisplay, workbookFormat(workbook, entriesStyle))
    worksheet.write('J{}'.format(count + 1), 'TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheet.merge_range('A{}:Y{}'.format(count + 2, count + 2), '', workbookFormat(workbook, topBorderStyle))

    for c in range(ord('K'), ord('Y') + 1):
            worksheet.write('{}{}'.format(chr(c), count + 1), "=SUM({}8:{}{})".format(chr(c), chr(c), count - 1),
                            workbookFormat(workbook, footerStyle))

    countcash = count + dfCashcount
    countecpay = count + dfCashcount + dfEcpaycount + 5
    countbc = count + dfCashcount + dfEcpaycount + dfBCcount + 10
    countbank = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + 15
    countcheck = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + 20

    paymentTypeWorksheet(worksheet, count, 'CASH TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countcash + 5, 'ECPAY TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countecpay + 5, 'BAYAD CENTER\nTYPE', workbookFormat(workbook, textWrapHeader))
    paymentTypeWorksheet(worksheet, countbc + 5, 'BANK TYPE', workbookFormat(workbook, headerStyle))
    paymentTypeWorksheet(worksheet, countbank + 5, 'CHECK TYPE', workbookFormat(workbook, headerStyle))

    dataframeStyle(worksheet, 'E', 'E', count + 6, countcash + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countcash + 11, count + countecpay + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countecpay + 11, countbc + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countbc + 6, count + countbank + 5, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'E', 'E', countbank + 11, countcheck + 5, workbookFormat(workbook, numFormat))

    sumPaymentType(worksheet, countcash, count, countcash, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countecpay, countcash + 5, countecpay, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countbc, countecpay + 5, countbc, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countbank, countbc + 5, countbank, workbookFormat(workbook, footerStyle))
    sumPaymentType(worksheet, countcheck, countbank + 5, countcheck, workbookFormat(workbook, footerStyle))

    totalPaymentType(worksheet, countcash, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countecpay, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countbc, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countbank, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))
    totalPaymentType(worksheet, countcheck, nodisplay, workbookFormat(workbook, docNameStyle), workbookFormat(workbook, entriesStyle), workbookFormat(workbook, topBorderStyle))

    counts = count + dfCashcount + dfEcpaycount + dfBCcount + dfBankcount + dfCheckcount + 5
    worksheet.merge_range('A{}:D{}'.format(counts + 24, counts + 24), 'DISBURSMENT', workbookFormat(workbook, headerStyle))
    worksheet.write('A{}'.format(counts + 25), '#', workbookFormat(workbook, headerStyle))
    worksheet.write('B{}'.format(counts + 25), 'DATE', workbookFormat(workbook, headerStyle))
    worksheet.write('C{}'.format(counts + 25), 'DESCRIPTION', workbookFormat(workbook, headerStyle))
    worksheet.write('D{}'.format(counts + 25), 'AMOUNT', workbookFormat(workbook, headerStyle))
    worksheet.merge_range('A{}:B{}'.format(counts + 27, counts + 27), 'TOTAL:', workbookFormat(workbook, docNameStyle))
    worksheet.write('D{}'.format(counts + 27), "=SUM(D{}:D{})".format(counts + 26, counts + 26), workbookFormat(workbook, borderFormatStyle))
    worksheet.merge_range('A{}:C{}'.format(counts + 29, counts + 29), 'NET COLLECTION:', workbookFormat(workbook, docNameStyle))
    worksheet.write('D{}'.format(counts + 29), "=K{}-D{}".format(count + 1, counts + 27), workbookFormat(workbook, borderFormatStyle))
    # worksheet.write('C{}'.format(count + count + 1), nodisplay, merge_format8)

    writer.close()

    # go back to the beginning of the stream
    output.seek(0)
    print('sending spreadsheet')
    filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
    return send_file(output, attachment_filename=filename, as_attachment=True)


@app.route("/dccr2", methods=['GET'])
# def get_data2():
#     output = 'test.xlsx'
#     dateStart = request.args.get('startDate')
#     dateEnd = request.args.get('endDate')
#     filename = "DCCR {}-{}.xlsx".format(dateStart, dateEnd)
#
#     payload = {'startDate': dateStart, 'endDate': dateEnd}
#     # url = "https://api360.zennerslab.com/Service1.svc/DCCRjson"
#     url = "https://rfc360-test.zennerslab.com/Service1.svc/DCCRjson"
#     r = requests.post(url, json=payload)
#     data_json = r.json()
#
#     writer = pd.ExcelWriter(filename, engine='xlsxwriter')
#     headers = ["Loan Account Number", "Customer Name", "Mobile Number", "OR Number", "OR Date", "Net Cash",
#                "Payment Source"]
#     df = pd.DataFrame(data_json['DCCRjsonResult'])
#     df = df[['loanAccountNo', 'customerName', 'mobileno', 'orNo', "postedDate", "amountApplied", "paymentSource"]]
#     df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)
#
#     workbook = writer.book
#     merge_format1 = workbook.add_format({'align': 'center'})
#     merge_format3 = workbook.add_format({'bold': True, 'align': 'center'})
#     xldate_header = "For the Period {} to {}".format(dateStart, dateEnd)
#
#     worksheet = writer.sheets["Sheet_1"]
#     worksheet.merge_range('A1:G1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheet.merge_range('A2:G2', 'RFC360 Kwikredit', merge_format1)
#     worksheet.merge_range('A3:G3', 'Daily Cash Collection Report', merge_format3)
#     worksheet.merge_range('A4:G4', xldate_header, merge_format1)
#
#     writer.save()
#
#     print('sending spreadsheet')
#     send_mail("cu.michaels@gmail.com", "jantzen@thegentlemanproject.com", "hello", "helloworld", filename,
#               'smtp.gmail.com', '587', 'cu.michaels@gmail.com', 'jantzen216')
#     return 'ok'
#     # return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/newmonthlyincome", methods=['GET'])
def get_monthly1():

    output = BytesIO()

    date = request.args.get('date')
    name = request.args.get('name')
    datetime_object = datetime.strptime(date, '%m/%d/%Y')
    month = datetime_object.strftime("%B")

    payload = {'date': date}
    url = "https://api360.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
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
        df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['appId', 'orDate'], inplace=True)
        df['orAmount'] = 0
        df["unappliedBalance"] = df['orAmount'] - (df['penaltyPaid'] + df['interestPaid'] + df['principalPaid'])
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'orDate')
        df = round(df, 2)
        df = df[['num', 'appId', 'loanAccountno', 'name', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
                 'orAmount', "orDate", "orNo"]]
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

@app.route("/monthlyincome", methods=['GET'])
# def get_monthly():
#
#     output = BytesIO()
#
#     date = request.args.get('date')
#     name = request.args.get('name')
#     datetime_object = datetime.strptime(date, '%m/%d/%Y')
#     month = datetime_object.strftime("%B")
#
#     payload = {'date': date}
#     url = "https://api360.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
#     url = "https://rfc360-test.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
#     r = requests.post(url, json=payload)
#     data_json = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["App ID", "Loan Account Number", "Customer Name", "Penalty Paid",
#                "Interest Paid", "Principal Paid", "Unapplied Balance", "Payment Amount", "OR Date", "OR Number"]
#     df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])
#
#     if df.empty:
#         count = df.shape[0] + 8
#         sumPenalty = 0
#         sumInterest = 0
#         sumPrincipal = 0
#         sumUnapplied = 0
#         total = 0
#         nodisplay = 'No Data'
#         df = pd.DataFrame(pd.np.empty((0, 10)))
#     else:
#         count = df.shape[0] + 8
#         nodisplay = ''
#         df['loanAccountno'] = df['loanAccountno'].map(lambda x: x.lstrip("'"))
#         df['appId'] = df['appId'].astype(int)
#         df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
#         df.sort_values(by=['appId'], inplace=True)
#         sumPenalty = pd.Series(df['penaltyPaid']).sum()
#         sumInterest = pd.Series(df['interestPaid']).sum()
#         sumPrincipal = pd.Series(df['principalPaid']).sum()
#         sumUnapplied = pd.Series(df['unappliedBalance']).sum()
#         total = pd.Series(df['paymentAmount']).sum()
#         df = df[['appId', 'loanAccountno', 'name', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
#                  'paymentAmount', "orDate", "orNo"]]
#     df.to_excel(writer, startrow=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)
#
#     workbook = writer.book
#     merge_format2 = workbook.add_format(docNameStyle)
#     merge_format4 = workbook.add_format(footerStyle)
#     merge_format6 = workbook.add_format(entriesStyle)
#     merge_format7 = workbook.add_format(headerStyle)
#     xldate_header = "For the month of {}".format(month)
#
#     worksheet = writer.sheets["Sheet_1"]
#
#     # list1 = [len(i) for i in headers]
#     # # list1 = np.array(headerlen)
#     #
#     # if df.empty:
#     #     list2 = list1
#     # else:
#     #     list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
#     #
#     # def function(list1, list2):
#     #     list3 = [max(value) for value in zip(list1, list2)]
#     #     return list3
#     #
#     # for col_num, value in enumerate(function(list1, list2)):
#     #     worksheet.set_column(col_num, col_num, value + 1)
#
#     range1 = 'G'
#     range2 = 'H'
#     range3 = 'J'
#     reportTitle = 'Mothly Income Report'
#     workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, reportTitle)
#
#     worksheet.merge_range('A6:A7', 'App ID', merge_format7)
#     worksheet.merge_range('B6:B7', "Loan Acct. #", merge_format7)
#     worksheet.merge_range('C6:C7', "Client's Name", merge_format7)
#     worksheet.merge_range('D6:D7', 'Penalty Paid', merge_format7)
#     worksheet.merge_range('E6:E7', 'Interest Paid', merge_format7)
#     worksheet.merge_range('F6:F7', 'Principal Paid', merge_format7)
#     worksheet.merge_range('G6:G7', 'Unapplied Balance', merge_format7)
#     worksheet.merge_range('H6:H7', 'Payment Amount', merge_format7)
#     worksheet.merge_range('I6:I7', 'OR Date', merge_format7)
#     worksheet.merge_range('J6:J7', 'OR #', merge_format7)
#
#     worksheet.merge_range('A{}:J{}'.format(count, count), nodisplay, merge_format6)
#     worksheet.merge_range('A{}:B{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)
#     worksheet.write('D{}'.format(count + 1), sumPenalty, merge_format4)
#     worksheet.write('E{}'.format(count + 1), sumInterest, merge_format4)
#     worksheet.write('F{}'.format(count + 1), sumPrincipal, merge_format4)
#     worksheet.write('G{}'.format(count + 1), sumUnapplied, merge_format4)
#     worksheet.write('H{}'.format(count + 1), total, merge_format4)
#
#     writer.close()
#
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Monthly Income {}.xlsx".format(date)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/monthlyincome2", methods=['GET'])
# def get_monthly2():
#
#     output = BytesIO()
#
#     date = request.args.get('date')
#     name = request.args.get('name')
#     datetime_object = datetime.strptime(date, '%m/%d/%Y')
#     month = datetime_object.strftime("%B")
#
#     payload = {'date': date}
#     # url = "https://api360.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
#     url = "https://rfc360-test.zennerslab.com/Service1.svc/monthlyIncomeReportJs"
#     r = requests.post(url, json=payload)
#     data_json = r.json()
#
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     headers = ["App ID", "Loan Account Number", "Customer Name", "Penalty Paid",
#                "Interest Paid", "Principal Paid", "Unapplied Balance", "Payment Amount"]
#     df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])
#
#     if df.empty:
#         count = df.shape[0] + 8
#         sumPenalty = 0
#         sumInterest = 0
#         sumPrincipal = 0
#         sumUnapplied = 0
#         total = 0
#         nodisplay = 'No Data'
#         df = pd.DataFrame(pd.np.empty((0, 8)))
#     else:
#         count = df.shape[0] + 8
#         nodisplay = ''
#         df['loanAccountno'] = df['loanAccountno'].map(lambda x: x.lstrip("'"))
#         df['appId'] = df['appId'].astype(int)
#         df["name"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
#         df.sort_values(by=['appId'], inplace=True)
#         sumPenalty = pd.Series(df['penaltyPaid']).sum()
#         sumInterest = pd.Series(df['interestPaid']).sum()
#         sumPrincipal = pd.Series(df['principalPaid']).sum()
#         sumUnapplied = pd.Series(df['unappliedBalance']).sum()
#         total = pd.Series(df['paymentAmount']).sum()
#         df = df[['appId', 'loanAccountno', 'name', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
#                  'paymentAmount']]
#     df.to_excel(writer, startrow=5, merge_cells=False, index=False, sheet_name="Sheet_1", header=headers)
#
#     workbook = writer.book
#     merge_format1 = workbook.add_format(periodStyle)
#     merge_format2 = workbook.add_format(docNameStyle)
#     merge_format3 = workbook.add_format(comNameStyle)
#     merge_format4 = workbook.add_format(footerStyle)
#     merge_format5 = workbook.add_format(generatedStyle)
#     merge_format6 = workbook.add_format(entriesStyle)
#     merge_format7 = workbook.add_format(headerStyle)
#     xldate_header = "For the month of {}".format(month)
#
#     worksheet = writer.sheets["Sheet_1"]
#
#     list1 = [len(i) for i in headers]
#     # list1 = np.array(headerlen)
#
#     if df.empty:
#         list2 = list1
#     else:
#         list2 = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
#
#     def function(list1, list2):
#         list3 = [max(value) for value in zip(list1, list2)]
#         return list3
#
#     for col_num, value in enumerate(function(list1, list2)):
#         worksheet.set_column(col_num, col_num, value + 1)
#
#     worksheet.merge_range('A1:H1', 'RADIOWEALTH FINANCE COMPANY, INC.', merge_format3)
#     worksheet.merge_range('A2:H2', 'RFC360 Kwikredit', merge_format1)
#     worksheet.merge_range('A3:H3', 'Monthly Income Report', merge_format3)
#     worksheet.merge_range('A4:H4', xldate_header, merge_format1)
#     worksheet.merge_range('A{}:H{}'.format(count - 1, count - 1), nodisplay, merge_format1)
#     worksheet.write('C{}'.format(count + 1), 'TOTAL', merge_format3)
#     worksheet.write('D{}'.format(count + 1), sumPenalty, merge_format4)
#     worksheet.write('E{}'.format(count + 1), sumInterest, merge_format4)
#     worksheet.write('F{}'.format(count + 1), sumPrincipal, merge_format4)
#     worksheet.write('G{}'.format(count + 1), sumUnapplied, merge_format4)
#     worksheet.write('H{}'.format(count + 1), total, merge_format4)
#     worksheet.merge_range('A{}:H{}'.format(count + 3, count + 3), 'Report Generated By :', merge_format2)
#     worksheet.merge_range('A{}:H{}'.format(count + 4, count + 5), name, merge_format2)
#     worksheet.merge_range('A{}:H{}'.format(count + 7, count + 7), 'Date & Time Report Generation ({})'.format(dateNow),
#                           merge_format2)
#
#     writer.close()
#
#     output.seek(0)
#     print('sending spreadsheet')
#     filename = "Monthly Income {}.xlsx".format(date)
#     return send_file(output, attachment_filename=filename, as_attachment=True)

@app.route("/booking", methods=['GET'])
def get_booking():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}
    url = "https://api360.zenners lab.com/Service1.svc/bookingReportJs"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/bookingReportJs"
    # url = "http://localhost:15021/Service1.svc/bookingReportJs"
    r = requests.post(url, json=payload)
    data_json = r.json()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "PRODUCT CODE", "SA", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "SUB PRODUCT", "PNV", "MLV", "FINANCE FEE", "HF",
               "DST", "NOTARIAL", "GCLI", "OMA", "TERM", "RATE", "MI", "APPLICATION DATE", "APPROVAL DATE", "BOOKING DATE", "FDD", "PROMO NAME"]
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
        df["customerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        dfDateFormat(df, 'forreleasingdate')
        dfDateFormat(df, 'approvalDate')
        dfDateFormat(df, 'generationDate')
        dfDateFormat(df, 'applicationDate')
        dfDateFormat(df, 'fdd')
        astype(df, 'loanId', int)
        astype(df, 'term', int)
        astype(df, 'actualRate', float)
        df.sort_values(by=['loanId'], inplace=True)
        count = df.shape[0] + 8
        df['num'] = numbers(df.shape[0])
        df = df[['num', 'channelName', 'partnerCode', 'outletCode', 'productCode', 'sa', 'loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "mlv", "insurance",
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
    dataframeStyle(worksheet, 'S', 'T', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'U', 'X', 8, count, workbookFormat(workbook, numFormat))
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
    url = "https://api360.zennerslab.com/Service1.svc/generateincentiveReportJSON"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/generateincentiveReportJSON"
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
        df["borrowerName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['agentName'], inplace=True)
        df['bookingDate'] = pd.to_datetime(df['bookingDate'])
        df['bookingDate'] = df['bookingDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'bookingDate')
        df = df[['num', 'bookingDate', 'loanId', 'borrowerName', 'refferalType', "SA", "dealerName", "loanType", "term",
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
    url = "https://api360.zennerslab.com/Service1.svc/maturedLoanReport"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/maturedLoanReport"
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
        df["fullName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        dfDateFormat(df, 'lastDueDate')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        df = df[['num', 'loanId', 'loanAccountNo', 'fullName', "mobileno", "term", "bMLV", "lastDueDate", "lastPayment",
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
    url = "https://api360.zennerslab.com/Service1.svc/dueTodayReport"
    # url = "https://rfc360-test.zennerslab.com/Service1.svc/dueTodayReport"
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
        df["fullName"] = df['firstName'] + ' ' + df['middleName'] + ' ' + df['lastName'] + ' ' + df['suffix']
        df.sort_values(by=['loanId'], inplace=True)
        dfDateFormat(df, 'monthlydue')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        df = df[["num", "loanId", "loanAccountNo", "fullName", "mobileno", "loanType", "term", "monthlyAmmortization",
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
    name = request.args.get('name')

    payload = {'loanId': loanId}

    url = "https://rfc360.mybluemix.net/customerLedger/ledgerByLoanId?loanId={}".format(loanId) #live
    # url = "https://rfc360-staging.mybluemix.net/customerLedger/ledgerByLoanId?loanId={}".format(loanId) #test
    r = requests.get(url, json=payload)

    ledgerData = requests.get(url).json()
    # print(ledgerData)

    data_json = {

	"borrowerDetails":
		[{
		 "appId":"1672",
		 "loanAccNum":"10101100999001300",
		 "borrowersName":"12345",
		 "collector":"678910",
		 "contactNum":"1112131415",
		 "address":"1617181920"
		}],

	"loanDetails":
		[{
		 "loanType":"2424",
		 "grossMI":"1957900",
		 "totalAddOnRate":"4500",
		 "terms":"36",
		 "disbursementDate":"978968",
		 "fdd":"42424"
		}],

	"collateralDetails":
		[{
		 "model":"242424",
		 "brand":"242424",
		 "serialNum":"242525",
		 "engineNum":"252552",
		 "plateNum":"226526254",
		 "orNum":"2352536"
		}],

	"acctStatDetails":
		[{
		 "expTerm":"21",
		 "remainingTerm":"15",
		 "miPaid":"23",
		 "monthsDue":"2",
		 "overdueAmount":"2618060"
        }],

    "obDetails":
        [{
		 "rfc":"27606000",
		 "penalty":"000",
		 "advances":"26180.60",
		 "total":"26722391",
		 "totalPayment":"43943549",
		 "lastPaymentDate":"253454534",
		}],

	"loanAccountSummary":
		[{
		 "total":"4545",
		 "paid":"34335",
		 "adj":"3535",
		 "billed":"35235",
		 "amountDue":"43553",
		 "balance":"535335",
		 "sBal":"322643"
		}]
}
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # headers = ["#", "DATE", "TRANS. TYPE", "BATCH CODE", "REF #", "BNK/ CHQ#", "PRIN", "INT", "INT ACCR", "PEN (5%)", "ADV", "TOTAL", "DUE", "OB"]
    headers = ["DATE", "TERM", "TRANSACTION TYPE", "PAYMENT TYPE", "REF NO", "CHECK #", "PENALTY INCUR", "PRINCIPAL",
               "INTEREST", "PENALTY PAID", "ADVANCES", "TOTAL", "DUE", "OB", "PAYMENT DATE", "OR NO", "OR DATE"]

    headersBorrower = ["APPLICATION ID", "LOAN ACCOUNT NO.", "BORROWER'S NAME", "COLLECTOR", "CONTACT NO.", "ADDRESS"]
    dataBorrower = ["appId", "loanAccNum", "borrowersName", "collector", "contactNum", "address"]
    headersLoan = ["LOAN TYPE", "GROSS MI", "TOTAL ADD-ON RATE", "TERMS", "DISBURSEMENT DATE", "FIRST DUE DATE"]
    dataLoan = ["loanType", "grossMI", "totalAddOnRate", "terms", "disbursementDate", "fdd"]
    headersCollateral = ["UNIT/MODEL/DESC.", "BRAND/MAKE", "SERIAL/CHASSIS NO.", "ENGINE NO.", "PLATE NO.", "O.R NO."]
    dataCollateral = ["model", "brand", "serialNum", "engineNum", "plateNum", "orNum"]
    headersAccStat = ["EXPIRED TERM", "REMAINING TERM", "NO. OF MI's PAID", "MONTHS DUE", "OVERDUE AMOUNT"]
    dataAcctStat = ["expTerm", "remainingTerm", "miPaid", "monthsDue", "overdueAmount"]
    headersOB = ["RFC", "PENALTY", "ADVANCES", "TOTAL", "TOTAL PAYMENT", "LAST PAYMENT DATE"]
    dataOB = ["rfc", "penalty", "advances", "total", "totalPayment", "lastPaymentDate"]
    headersLoanSummary = ["TOTAL", "PAID", "ADJ", "BILLED", "AMT DUE", "BAL.", "SHOULD BE BAL."]
    hdataLoanSummary = ["TOTAL", "PAID", "ADJ", "BILLED", "AMT DUE", "BAL.", "SHOULD BE BAL."]

    dfLedger = pd.DataFrame(ledgerData['data']['transactions'])
    dfBorrower = pd.DataFrame(data_json['borrowerDetails'])
    dfLoan = pd.DataFrame(data_json['loanDetails'])
    dfCollateral = pd.DataFrame(data_json['collateralDetails'])
    dfAcctStat = pd.DataFrame(data_json['acctStatDetails'])
    dfOBDetails = pd.DataFrame(data_json['obDetails'])
    dfLoanSummary = pd.DataFrame(data_json['loanAccountSummary'])

    list1 = [len(i) for i in headers]

    if dfLedger.empty:
        count = dfLedger.shape[0] + 33
        nodisplay = 'No Data'
        dfLedger = pd.DataFrame(pd.np.empty((0, 12)))
        list2 = list1
    else:
        count = dfLedger.shape[0] + 33
        dfLedger['orDate'] = dfLedger['orDate'].loc[dfLedger['orDate'].str.contains("/")]
        dfLedger['paymentDate'] = dfLedger['paymentDate'].loc[dfLedger['paymentDate'].str.contains("/")]
        dfDateFormat(dfLedger, 'orDate')
        dfDateFormat(dfLedger, 'paymentDate')
        dfDateFormat(dfLedger, 'date')
        astype(dfLedger, 'penaltyIncur', float)
        astype(dfLedger, 'principal', float)
        astype(dfLedger, 'interest', float)
        astype(dfLedger, 'penaltyPaid', float)
        astype(dfLedger, 'advances', float)
        astype(dfLedger, 'mi', float)
        astype(dfLedger, 'amountDue', float)
        astype(dfLedger, 'ob', float)
        dfLedger = dfLedger[["date", "term", "type", "paymentType", "refNo", "checkNo", "penaltyIncur", "principal", "interest", "penaltyPaid",
             "advances", "mi", "amountDue", "ob", "paymentDate", "orNo", "orDate"]]
        list2 = [max([len(str(s)) for s in dfLedger[col].values]) for col in dfLedger.columns]

    # dfLoanSummary = dfLoanSummary.style.set_properties(**styles)
    # dfLedger = dfLedger.style.set_properties(**styles)
    dfLoanSummary.to_excel(writer, startrow=17, startcol=7, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)
    dfLedger.to_excel(writer, startrow=33, merge_cells=False, index=False, sheet_name="Sheet_1", header=None)

    workbook = writer.book

    worksheet = writer.sheets["Sheet_1"]

    dataframeStyle(worksheet, 'A', 'B', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'C', 'D', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'C', 'D', 8, count, workbookFormat(workbook, stringFormat))
    dataframeStyle(worksheet, 'E', 'F', 8, count, workbookFormat(workbook, defaultFormat))
    dataframeStyle(worksheet, 'G', 'O', 8, count, workbookFormat(workbook, numFormat))
    dataframeStyle(worksheet, 'P', 'Q', 8, count, workbookFormat(workbook, defaultFormat))

    for col_num, value in enumerate(columnWidth(list1, list2)):
        worksheet.set_column(col_num, col_num, value)

    range1 = 'N'
    range2 = 'O'
    range3 = 'Q'
    companyName = 'RFSC'
    reportTitle = 'CUSTOMER LEDGER'
    xldate_header = 'Loan ID : {}'.format(loanId)
    branchName = 'Nationwide'

    workSheet(workbook, worksheet, range1, range2, range3, xldate_header, name, companyName, reportTitle, branchName)

    def numbers(numRange, addnum):
        number = [number + addnum for number in range(numRange)]
        return number

    def alphabetRange(firstRange, secondRange):
        alphaList = [chr(c) for c in range(ord(firstRange), ord(secondRange) + 1)]
        return alphaList

    worksheet.merge_range('A6:F6', 'BORROWER DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(numbers(6, 7), headersBorrower):
        worksheet.merge_range('A{}:C{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(numbers(6, 7), dataBorrower):
        worksheet.merge_range('D{}:E{}'.format(x, x), dfBorrower[y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('G6:J6', 'LOAN DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(numbers(6, 7), headersLoan):
        worksheet.merge_range('G{}:H{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(numbers(6, 7), dataLoan):
        worksheet.merge_range('I{}:J{}'.format(x, x), dfLoan[y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('L6:O6', 'COLLATERAL DETAILS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(numbers(6, 7), headersCollateral):
        worksheet.merge_range('L{}:M{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(numbers(6, 7), dataCollateral):
        worksheet.merge_range('N{}:O{}'.format(x, x), dfCollateral[y], workbookFormat(workbook, ledgerDataStyle))

    worksheet.merge_range('A15:C15', 'ACCOUNT STATUS', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(numbers(5, 16), headersAccStat):
        worksheet.merge_range('A{}:C{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    worksheet.merge_range('D15:E15', 'NORMAL', workbookFormat(workbook, undStyle))

    for x, y in zip(numbers(5, 16), dataAcctStat):
        worksheet.merge_range('D{}:E{}'.format(x, x), dfAcctStat[y], workbookFormat(workbook, defaultFormat))

    worksheet.merge_range('H15:N15', 'LOAN ACCOUNT SUMMARY', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(alphabetRange('H', 'N'), headersLoanSummary):
        worksheet.write('{}17'.format(x), '{}'.format(y)
                        , workbookFormat(workbook, sumStyle))

    worksheet.write('G18', 'GRAND TOTAL', workbookFormat(workbook, stringFormat))
    worksheet.write('G20', 'TOTAL PNV', workbookFormat(workbook, stringFormat))
    worksheet.write('G21', 'RFC', workbookFormat(workbook, stringFormat))
    worksheet.write('G22', 'PRINCIPAL', workbookFormat(workbook, defaultFormat))
    worksheet.write('G23', 'INTEREST', workbookFormat(workbook, defaultFormat))
    worksheet.write('G25', 'PENALTY', workbookFormat(workbook, stringFormat))
    worksheet.write('G26', 'ADVANCES', workbookFormat(workbook, stringFormat))

    worksheet.merge_range('A22:F22', 'OUTSTANDING BALANCE:', workbookFormat(workbook, ledgerStyle))
    for x, y in zip(numbers(6, 23), headersOB):
        if (y == 'TOTAL PAYMENT'):
            worksheet.merge_range('A28:C28'.format(x, x), 'TOTAL PAYMENT', workbookFormat(workbook, ledgerNameStyle))
        elif(y == 'LAST PAYMENT DATE'):
            worksheet.merge_range('A30:C30'.format(x, x), 'LAST PAYMENT DATE', workbookFormat(workbook, ledgerNameStyle))
        else:
            worksheet.merge_range('A{}:C{}'.format(x, x), '{}'.format(y), workbookFormat(workbook, ledgerNameStyle))

    for x, y in zip(numbers(6, 23), dataOB):
        if(x == 25):
            worksheet.write('E25', dfOBDetails[y], workbookFormat(workbook, defaultUnderlineFormat))
        elif(x == 26):
            worksheet.write('E26', dfOBDetails[y], workbookFormat(workbook, defaultUnderlineFormat))
        elif(y == 'totalPayment'):
            worksheet.merge_range('D28:E28'.format(x, x), dfOBDetails['totalPayment'], workbookFormat(workbook, defaultFormat))
        elif(y == 'lastPaymentDate'):
            worksheet.merge_range('D30:E30'.format(x, x), dfOBDetails[y], workbookFormat(workbook, defaultFormat))
        else:
            worksheet.merge_range('D{}:E{}'.format(x, x), dfOBDetails['lastPaymentDate'], workbookFormat(workbook, defaultFormat))

    headersList = [i for i in headers]

    for x, y in zip(alphabet(range3), headersList):
        worksheet.merge_range('{}32:{}33'.format(x, x), '{}'.format(y), workbookFormat(workbook, headerStyle))

    # worksheet.merge_range('A{}:Q{}'.format(count, count), '', merge_format3)
    # worksheet.merge_range('A{}:C{}'.format(count + 1, count + 1), 'GRAND TOTAL:', merge_format2)
    #
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
