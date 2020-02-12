from flask import Flask, Blueprint, request, jsonify, send_file, render_template, make_response
import json
import requests
import pandas as pd
import numpy as np
import openpyxl
import flask_excel as excel
from io import BytesIO, StringIO
import os
import subprocess
import platform


from datetime import date


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

import jinja2
import pdfkit


fmtDate = "%m/%d/%y"
fmtTime = "%I:%M %p"
now_utc = datetime.now(timezone('UTC'))
now_pacific = now_utc.astimezone(timezone('Asia/Manila'))

dateNow = now_pacific.strftime(fmtDate)
timeNow = now_pacific.strftime(fmtTime)

pdf_api = Blueprint('pdf_api', __name__)

# serviceUrl = "https://api360.zennerslab.com/Service1.svc/{}" #rfc-service-live
lambdaUrl = "https://ia-lambda-live.mybluemix.net/{}" #lambda-bluemix-live
bluemixUrl = "https://rfc360.mybluemix.net/{}" #rfc-bluemix-live
serviceUrl = "http://localhost:15021/Service1.svc/{}" #rfc-localhost

def _get_pdfkit_config():
     """wkhtmltopdf lives and functions differently depending on Windows or Linux. We
      need to support both since we develop on windows but deploy on Heroku.

     Returns:
         A pdfkit configuration
     """
     if platform.system() == 'Windows':
         return pdfkit.configuration(wkhtmltopdf=os.environ.get('WKHTMLTOPDF_BINARY', 'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'))
     else:
         WKHTMLTOPDF_CMD = subprocess.Popen(['which', os.environ.get('WKHTMLTOPDF_BINARY', 'wkhtmltopdf-pack')], stdout=subprocess.PIPE).communicate()[0].decode('utf-8').strip()
         print(WKHTMLTOPDF_CMD)
         return pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_CMD)

def numbers(numRange):
    number = [number + 1 for number in range(numRange)]
    return number

def astype(df, colName, type):
    df[colName] = df[colName].astype(type)
    return df[colName]

def dfDateFormat(df, colDateName):
    df[colDateName] = pd.to_datetime(df[colDateName])
    df[colDateName] = df[colDateName].map(lambda x: x.strftime('%m/%d/%y') if pd.notnull(x) else '')
    return df[colDateName]

def startDateFormat(dateStart):
    dateStart_object = datetime.strptime(dateStart, '%m/%d/%Y')
    payloaddateStart = dateStart_object.strftime('%m/%d/%y')
    return payloaddateStart

def endDateFormat(dateEnd):
    dateStart_object = datetime.strptime(dateEnd, '%m/%d/%Y')
    payloaddateEnd= dateStart_object.strftime('%m/%d/%y')
    return payloaddateEnd

def numberFormat(dfName):
    format = "{0:,.2f}".format(round(pd.Series(dfName).sum(), 2))
    return format

def dfNumberFormat(dfName):
    dfName = pd.to_numeric(dfName.fillna(0), errors='coerce')
    dfName = dfName.map('{:,.2f}'.format)
    return dfName

@pdf_api.route("/")
def pdfList():
    return "list of pdfs"


@pdf_api.route("/collectionreport", methods=['GET'])
def collection_pdf():
    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("collection")
    r = requests.post(url, json=payload)
    data = r.json()
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "AMT DUE", "FDD", "PNV", "MLV", "MI", "TERM", "PEN", "INT",
               "PRIN", "UNPAID MOS", "PAID MOS"]
    df = pd.DataFrame(data['collectionResult'])
    list1 = [len(i) for i in headers]
    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 21)))
        sumamountDues = ''
        sumpnv = ''
        summlv = ''
        summi = ''
        sumsumOfPenalty = ''
        sumtotalInterest = ''
        sumtotalPrincipal = ''
        sumhf = ''
        sumdst = ''
        sumnotarial = ''
        sumgcli = ''
        sumoutstandingBalance = ''
        sumtotalPayment = ''
    else:
        astype(df, 'dd', int)
        astype(df, 'term', int)
        astype(df, 'unapaidMonths', int)
        astype(df, 'paidMonths', int)
        astype(df, 'loanIndex', int)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['hf'] = 0
        df['dst'] = 0
        df['notarial'] = 0
        df['gcli'] = 0
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'fdd')
        df = round(df, 2)
        sumamountDues = numberFormat(df['amountDue'])
        sumpnv = numberFormat(df['pnv'])
        summlv = numberFormat(df['mlv'])
        summi = numberFormat(df['mi'])
        sumsumOfPenalty = numberFormat(df['sumOfPenalty'])
        sumtotalInterest = numberFormat(df['totalInterest'])
        sumtotalPrincipal = numberFormat(df['totalPrincipal'])
        sumhf = numberFormat(df['hf'])
        sumdst = numberFormat(df['dst'])
        sumnotarial = numberFormat(df['notarial'])
        sumgcli = numberFormat(df['gcli'])
        sumoutstandingBalance = numberFormat(df['outstandingBalance'])
        sumtotalPayment = numberFormat(df['totalPayment'])
        df = df[["num","loanId", "loanAccountNo", "name", "amountDue", "fdd", "pnv", "mlv", "mi", "term",
                 "sumOfPenalty", "totalInterest", "totalPrincipal", "unapaidMonths", "paidMonths", "hf", "dst", "notarial", "gcli", "outstandingBalance", "status",
                 "totalPayment"]]

    # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 37)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)

        df50['num'] = df['num'].map('{:.0f}'.format)
        # df50['loanId'] = df['loanId'].map('{:.0f}'.format)
        df50['term'] = df['term'].map('{:.0f}'.format)
        df50['unapaidMonths'] = df['unapaidMonths'].map('{:.0f}'.format)
        df50['paidMonths'] = df['paidMonths'].map('{:.0f}'.format)
        df50['amountDue'] = dfNumberFormat(df50['amountDue'])
        df50['pnv'] = dfNumberFormat(df50['pnv'])
        df50['mlv'] = dfNumberFormat(df50['mlv'])
        df50['mi'] = dfNumberFormat(df50['mi'])
        df50['sumOfPenalty'] = dfNumberFormat(df50['sumOfPenalty'])
        df50['totalInterest'] = dfNumberFormat(df50['totalInterest'])
        df50['totalPrincipal'] = dfNumberFormat(df50['totalPrincipal'])
        df50['hf'] = dfNumberFormat(df50['hf'])
        df50['dst'] = dfNumberFormat(df50['dst'])
        df50['notarial'] = dfNumberFormat(df50['notarial'])
        df50['gcli'] = dfNumberFormat(df50['gcli'])
        df50['outstandingBalance'] = dfNumberFormat(df50['outstandingBalance'])
        df50['totalPayment'] = dfNumberFormat(df50['totalPayment'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'loanId'] = 'SUB TOTAL:'
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'unapaidMonths'] = ''
        df50.loc['Total', 'paidMonths'] = ''

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('collection_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"), range=xldate_header, time=timeNow,
                           name=name, df=split_df_to_chunks_of_50, sumamountDues=sumamountDues,
                           sumpnv=sumpnv, summlv=summlv, summi=summi, sumsumOfPenalty=sumsumOfPenalty, sumtotalInterest=sumtotalInterest,
                           sumtotalPrincipal=sumtotalPrincipal, sumhf=sumhf, sumdst=sumdst, sumnotarial=sumnotarial, sumgcli=sumgcli,
                           sumoutstandingBalance=sumoutstandingBalance, sumtotalPayment=sumtotalPayment, empty_df=df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Collection Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response


@pdf_api.route("/dccr", methods=['GET'])
def dccr_pdf():

    output = BytesIO()

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("DCCRjsonNew")

    r = requests.post(url, json=payload)
    data_json = r.json()
    headers1 = ["#", "PAYMENT", "LOAN ACCT. #", "CUSTOMER NAME", "OR DATE", "OR NUM", "BANK", "CHECK #", "PAYMENT DATE"]
    df = pd.DataFrame(data_json['DCCRjsonNewResult'])
    df1 = pd.DataFrame(data_json['DCCRjsonNewResult']).copy()

    if df.empty or df1.empty:
        df = pd.DataFrame(pd.np.empty((0, 25)))
        df1['num1'] = ''
        dfCash = pd.DataFrame(pd.np.empty((0, 25)))
        dfEcpay = pd.DataFrame(pd.np.empty((0, 25)))
        dfBC = pd.DataFrame(pd.np.empty((0, 25)))
        dfBank = pd.DataFrame(pd.np.empty((0, 25)))
        dfCheck = pd.DataFrame(pd.np.empty((0, 25)))
        dfGPRS = pd.DataFrame(pd.np.empty((0, 25)))
        df2 = pd.DataFrame(pd.np.empty((0, 25)))
        sumamount = ''
        sumcash = ''
        sumpaymentCheck = ''
        sumpaidPrincipal = ''
        sumpaidInterest = ''
        sumadvances = ''
        sumpaidPenalty = ''
        dfCashTotal = ''
        dfCashAmount = ''
        dfCashCheck = ''
        dfEcpayTotal = ''
        dfEcpayAmount = ''
        dfEcpayCheck = ''
        dfBCTotal = ''
        dfBCAmount = ''
        dfBCCheck = ''
        dfBankTotal = ''
        dfBankAmount = ''
        dfBankCheck = ''
        dfCheckTotal = ''
        dfCheckAmount = ''
        dfCheckCheck = ''
        dfGPRSTotal = ''
        dfGPRSAmount = ''
        dfGPRSCheck = ''
    else:
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
        sumamount = numberFormat(df['amount'])
        sumcash = numberFormat(df['cash'])
        sumpaymentCheck = numberFormat(df['paymentCheck'])
        sumpaidPrincipal = numberFormat(df['paidPrincipal'])
        sumpaidInterest = numberFormat(df['paidInterest'])
        sumadvances = numberFormat(df['advances'])
        sumpaidPenalty = numberFormat(df['paidPenalty'])

        dfCashTotal =  numberFormat(dfCash['total'])
        dfCashAmount =  numberFormat(dfCash['amount'])
        dfCashCheck =  numberFormat(dfCash['paymentCheck'])

        dfEcpayTotal = numberFormat(dfEcpay['total'])
        dfEcpayAmount = numberFormat(dfEcpay['amount'])
        dfEcpayCheck = numberFormat(dfEcpay['paymentCheck'])

        dfBCTotal = numberFormat(dfBC['total'])
        dfBCAmount = numberFormat(dfBC['amount'])
        dfBCCheck = numberFormat(dfBC['paymentCheck'])

        dfBankTotal = numberFormat(dfBank['total'])
        dfBankAmount = numberFormat(dfBank['amount'])
        dfBankCheck = numberFormat(dfBank['paymentCheck'])

        dfCheckTotal = numberFormat(dfCheck['total'])
        dfCheckAmount = numberFormat(dfCheck['amount'])
        dfCheckCheck = numberFormat(dfCheck['paymentCheck'])

        dfGPRSTotal = numberFormat(dfGPRS['total'])
        dfGPRSAmount = numberFormat(dfGPRS['amount'])
        dfGPRSCheck = numberFormat(dfGPRS['paymentCheck'])

        df = df[['num', 'transType', 'loanAccountNo', 'newCustomerName', 'orDate', 'orNo', 'bank', 'checkNo', 'date',
                 'amount', 'cash', 'paymentCheck', 'paidPrincipal', 'paidInterest', 'advances', 'paidPenalty']]
        dfCash = dfCash[['dfCashnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfEcpay = dfEcpay[['dfEcpaynum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBC = dfBC[['dfBCnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfBank = dfBank[['dfBanknum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        dfCheck = dfCheck[['dfChecknum', 'orDate', 'orNo', 'transType', 'amount', 'total', 'paymentCheck']]
        dfGPRS = dfGPRS[['dfGPRSnum', 'orDate', 'orNo', 'paymentSource', 'total', 'amount', 'paymentCheck']]
        df2 = df1[['num1', 'num1', 'num1', 'num1']]

    # split the dataframe into rows of 50
    split_df = split_dataframe_to_chunks(df, 37)
    split_dfCash = split_dataframe_to_chunks(dfCash, 42)
    split_dfEcpay = split_dataframe_to_chunks(dfEcpay, 42)
    split_dfBC = split_dataframe_to_chunks(dfBC, 42)
    split_dfBank = split_dataframe_to_chunks(dfBank, 42)
    split_dfCheck = split_dataframe_to_chunks(dfCheck, 42)
    split_dfGPRS = split_dataframe_to_chunks(dfGPRS, 42)
    split_df2 = split_dataframe_to_chunks(df2, 42)

    # add Totals row to each dataframe
    for df1 in split_df:
        df1.loc['Total'] = round(df1.select_dtypes(pd.np.number).sum(), 2)
        df1['num'] = df['num'].map('{:.0f}'.format)
        df1['amount'] = dfNumberFormat(df1['amount'])
        df1['cash'] = dfNumberFormat(df1['cash'])
        df1['paymentCheck'] = dfNumberFormat(df1['paymentCheck'])
        df1['paidPrincipal'] = dfNumberFormat(df1['paidPrincipal'])
        df1['paidInterest'] = dfNumberFormat(df1['paidInterest'])
        df1['advances'] = dfNumberFormat(df1['advances'])
        df1['paidPenalty'] = dfNumberFormat(df1['paidPenalty'])
        df1.loc['Total'] = df1.loc['Total'].replace(np.nan, '', regex=True)
        df1.loc['Total', 'num'] = ''
        df1.loc['Total', 'transType'] = 'SUB TOTAL:'

    for df2 in split_dfCash:
        df2.loc['Total'] = round(df1.select_dtypes(pd.np.number).sum(), 2)
        df2['dfCashnum'] = dfCash['dfCashnum'].map('{:.0f}'.format)
        df2['total'] = dfNumberFormat(df2['total'])
        df2['amount'] = dfNumberFormat(df2['amount'])
        df2['paymentCheck'] = dfNumberFormat(df2['paymentCheck'])
        df2.loc['Total'] = df2.loc['Total'].replace(np.nan, '', regex=True)
        df2.loc['Total', 'dfCashnum'] = 'SUB TOTAL:'

    for df3 in split_dfEcpay:
        df3.loc['Total'] = round(df3.select_dtypes(pd.np.number).sum(), 2)
        df3['dfEcpaynum'] = dfEcpay['dfEcpaynum'].map('{:.0f}'.format)
        df3['total'] = dfNumberFormat(df3['total'])
        df3['amount'] = dfNumberFormat(df3['amount'])
        df3['paymentCheck'] = dfNumberFormat(df3['paymentCheck'])
        df3.loc['Total'] = df3.loc['Total'].replace(np.nan, '', regex=True)
        df3.loc['Total', 'dfEcpaynum'] = 'SUB TOTAL:'
    for df4 in split_dfBC:
        df4.loc['Total'] = round(df4.select_dtypes(pd.np.number).sum(), 2)
        df4['dfBCnum'] = dfBC['dfBCnum'].map('{:.0f}'.format)
        df4['total'] = dfNumberFormat(df4['total'])
        df4['amount'] = dfNumberFormat(df4['amount'])
        df4['paymentCheck'] = dfNumberFormat(df4['paymentCheck'])
        df4.loc['Total'] = df4.loc['Total'].replace(np.nan, '', regex=True)
        df4.loc['Total', 'dfBCnum'] = 'SUB TOTAL:'
    for df5 in split_dfBank:
        df5.loc['Total'] = round(df5.select_dtypes(pd.np.number).sum(), 2)
        df5['dfBanknum'] = dfBank['dfBanknum'].map('{:.0f}'.format)
        df5['total'] = dfNumberFormat(df5['total'])
        df5['amount'] = dfNumberFormat(df5['amount'])
        df5['paymentCheck'] = dfNumberFormat(df5['paymentCheck'])
        df5.loc['Total'] = df5.loc['Total'].replace(np.nan, '', regex=True)
        df5.loc['Total', 'dfBanknum'] = 'SUB TOTAL:'
    for df6 in split_dfCheck:
        df6.loc['Total'] = round(df6.select_dtypes(pd.np.number).sum(), 2)
        df6['dfChecknum'] = dfCheck['dfChecknum'].map('{:.0f}'.format)
        df6['total'] = dfNumberFormat(df6['total'])
        df6['amount'] = dfNumberFormat(df6['amount'])
        df6['paymentCheck'] = dfNumberFormat(df6['paymentCheck'])
        df6.loc['Total'] = df6.loc['Total'].replace(np.nan, '', regex=True)
        df6.loc['Total', 'dfChecknum'] = 'SUB TOTAL:'
    for df7 in split_dfGPRS:
        df7.loc['Total'] = round(df7.select_dtypes(pd.np.number).sum(), 2)
        df7['dfGPRSnum'] = dfGPRS['dfGPRSnum'].map('{:.0f}'.format)
        df7['total'] = dfNumberFormat(df7['total'])
        df7['amount'] = dfNumberFormat(df7['amount'])
        df7['paymentCheck'] = dfNumberFormat(df7['paymentCheck'])
        df7.loc['Total'] = df7.loc['Total'].replace(np.nan, '', regex=True)
        df7.loc['Total', 'dfGPRSnum'] = 'SUB TOTAL:'
    for df8 in split_df2:
        df8.loc['Total'] = round(df8.select_dtypes(pd.np.number).sum(), 2)
        df8.loc['Total'] = df8.loc['Total'].replace(np.nan, '', regex=True)
        df8.loc['Total', 'num'] = 'SUB TOTAL:'

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('dccr_template.html', headers=headers1, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                           name=name, df=split_df, split_dfCash=split_dfCash, split_dfEcpay=split_dfEcpay, split_dfBC=split_dfBC,
                           split_dfBank=split_dfBank, split_dfCheck=split_dfCheck, split_dfGPRS=split_dfGPRS, split_df2=split_df2,
                           sumamount=sumamount, sumcash=sumcash, sumpaymentCheck=sumpaymentCheck, sumpaidPrincipal=sumpaidPrincipal,
                           sumpaidInterest=sumpaidInterest, sumadvances=sumadvances, sumpaidPenalty=sumpaidPenalty, empty_df=df.empty,
                           empty_dfCash=dfCash.empty, empty_dfEcpay=dfEcpay.empty, empty_dfBC=dfBC.empty, empty_dfBank=dfBank.empty,
                           empty_dfCheck=dfCheck.empty, empty_dfGPRS=dfGPRS.empty,
                           dfCashTotal=dfCashTotal, dfCashAmount=dfCashAmount, dfCashCheck=dfCashCheck, dfEcpayTotal=dfEcpayTotal,
                           dfEcpayAmount=dfEcpayAmount, dfEcpayCheck=dfEcpayCheck, dfBCTotal=dfBCTotal, dfBCAmount=dfBCAmount,
                           dfBCCheck=dfBCCheck, dfBankTotal=dfBankTotal, dfBankAmount=dfBankAmount, dfBankCheck=dfBankCheck,
                           dfCheckTotal=dfCheckTotal, dfCheckAmount=dfCheckAmount, dfCheckCheck=dfCheckCheck, dfGPRSTotal=dfGPRSTotal,
                           dfGPRSAmount=dfGPRSAmount, dfGPRSCheck=dfGPRSCheck
                           )
    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=DCCR {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/accountingAgingReport", methods=['GET'])
def aging_pdf():
    output = BytesIO()

    name = request.args.get('name')
    dates = request.args.get('date')

    payload = {'date': dates}

    url = lambdaUrl.format("reports/accountingAgingReport")
    r = requests.post(url, json=payload)
    data = r.json()

    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "APP ID", "LOAN ACCT #", "CUSTOMER NAME",
               "COLLECTOR", "FDD", "LAST PAID DATE", "TERM", "EXP TERM", "MI", "STAT", "OUTS BAL.", "BMLV", "BUCKET",
               "CURR. TODAY"]
    agingp1DF = pd.DataFrame(data)


    if agingp1DF.empty:
        agingp1DF = pd.DataFrame(pd.np.empty((0, 19)))
        summonthlyInstallment = ''
        sumob = ''
        sumrunningMLV = ''
        sumtoday = ''
        sum1 = ''
        sum31 = ''
        sum61 = ''
        sum91 = ''
        sum121 = ''
        sum151 = ''
        sum181 = ''
        sum360 = ''
        sumtotal = ''
        sumduePrincipal = ''
        dueInterest = ''
        sumduePenalty = ''
        sumamountSum = ''
    else:
        agingp1DF['num'] = numbers(agingp1DF.shape[0])
        astype(agingp1DF, 'term', int)
        astype(agingp1DF, 'expiredTerm', int)
        # astype(agingp1DF, 'appId', int)
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
        agingp1DF["newCustomerName"] = agingp1DF['lastName'] + ', ' + agingp1DF['firstName'] + ' ' + agingp1DF[
            'middleName'] + ' ' + agingp1DF['suffix']
        agingp1DF['ob'] = agingp1DF['notDue'] + agingp1DF['monthDue']
        agingp1DF = round(agingp1DF, 2)
        summonthlyInstallment = numberFormat(agingp1DF['monthlyInstallment'])
        sumob = numberFormat(agingp1DF['ob'])
        sumrunningMLV = numberFormat(agingp1DF['runningMLV'])
        sumtoday = numberFormat(agingp1DF['today'])
        sum1 = numberFormat(agingp1DF['1-30'])
        sum31 = numberFormat(agingp1DF['31-60'])
        sum61 = numberFormat(agingp1DF['61-90'])
        sum91 = numberFormat(agingp1DF['91-120'])
        sum121 = numberFormat(agingp1DF['121-150'])
        sum151 = numberFormat(agingp1DF['151-180'])
        sum181 = numberFormat(agingp1DF['181-360'])
        sum360 = numberFormat(agingp1DF['360 & over'])
        sumtotal = numberFormat(agingp1DF['total'])
        sumduePrincipal = numberFormat(agingp1DF['duePrincipal'])
        dueInterest = numberFormat(agingp1DF['dueInterest'])
        sumduePenalty = numberFormat(agingp1DF['duePenalty'])
        sumamountSum = numberFormat(agingp1DF['amountSum'])
        agingp1DF = agingp1DF[
            ["num", "channelName", "partnerCode", "outletCode", "appId", "loanAccountNumber", "newCustomerName",
             "alias", "fdd", "lastPaymentDate", "term", "expiredTerm", "monthlyInstallment", "stats", "ob",
             "runningMLV", "bucketing", "today",
             "1-30", "31-60", "61-90", "91-120", "121-150", "151-180", "181-360", "360 & over", "total", "duePrincipal",
             "dueInterest", "duePenalty", "amountSum"]]

   # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(agingp1DF, 38)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['num'] = agingp1DF['num'].map('{:.0f}'.format)
        # df50['appId'] = agingp1DF['appId'].map('{:.0f}'.format)
        df50['term'] = agingp1DF['term'].map('{:.0f}'.format)
        df50['expiredTerm'] = agingp1DF['expiredTerm'].map('{:.0f}'.format)
        df50['monthlyInstallment'] = dfNumberFormat(df50['monthlyInstallment'])
        df50['ob'] = dfNumberFormat(df50['ob'])
        df50['runningMLV'] = dfNumberFormat(df50['runningMLV'])
        df50['today'] = dfNumberFormat(df50['today'])
        df50['today'] = dfNumberFormat(df50['today'])
        df50['today'] = dfNumberFormat(df50['today'])
        df50['1-30'] = dfNumberFormat(df50['1-30'])
        df50['31-60'] = dfNumberFormat(df50['31-60'])
        df50['61-90'] = dfNumberFormat(df50['61-90'])
        df50['91-120'] = dfNumberFormat(df50['91-120'])
        df50['121-150'] = dfNumberFormat(df50['121-150'])
        df50['151-180'] = dfNumberFormat(df50['151-180'])
        df50['181-360'] = dfNumberFormat(df50['181-360'])
        df50['360 & over'] = dfNumberFormat(df50['360 & over'])
        df50['total'] = dfNumberFormat(df50['total'])
        df50['duePrincipal'] = dfNumberFormat(df50['duePrincipal'])
        df50['dueInterest'] = dfNumberFormat(df50['dueInterest'])
        df50['duePenalty'] = dfNumberFormat(df50['duePenalty'])
        df50['amountSum'] = dfNumberFormat(df50['amountSum'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'channelName'] = 'SUB TOTAL:'
        df50.loc['Total', 'appId'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'expiredTerm'] = ''

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('aging_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"), range=xldate_header,
                           time=timeNow,name=name, df=split_df_to_chunks_of_50, summonthlyInstallment=summonthlyInstallment,
                           sumob=sumob, sumrunningMLV=sumrunningMLV, sumtoday=sumtoday, sum1=sum1, sum31=sum31, sum61=sum61,
                           sum91=sum91, sum121=sum121, sum151=sum151, sum181=sum181, sum360=sum360, sumtotal=sumtotal,
                           sumduePrincipal=sumduePrincipal, dueInterest=dueInterest, sumduePenalty=sumduePenalty, sumamountSum=sumamountSum,
                           empty_df=agingp1DF.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Aging Report as of {}.pdf'.format(dates)

    return response

@pdf_api.route("/booking", methods=['GET'])
def bookingPDF():

    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("bookingReportJs")

    r = requests.post(url, json=payload)
    data_json = r.json()

    headers = ["#", "CHANNEL NAME", "PARTNER CODE", "OUTLET CODE", "PRODUCT CODE", "SA", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "SUB PRODUCT", "PNV", "MLV", "FINANCE FEE", "HF",
               "DST", "NF", "GCLI", "OMA", "TERM (MOS)", "RATE(%)", "MI", "APPLICATION DATE", "APPROVAL DATE", "BOOKING DATE", "FDD", "PROMO NAME"]
    df = pd.DataFrame(data_json['bookingReportJsResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 23)))
        pnvsum = ''
        mlvsum = ''
        interestsum = ''
        handlingFeesum = ''
        dstsum = ''
        notarialsum = ''
        gclisum = ''
        otherFeessum = ''
        monthlyAmountsum = ''
    else:
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

        pnvsum = numberFormat(df['PNV'])
        mlvsum = numberFormat(df['mlv'])
        interestsum = numberFormat(df['interest'])
        handlingFeesum = numberFormat(df['handlingFee'])
        dstsum = numberFormat(df['dst'])
        notarialsum = numberFormat(df['notarial'])
        gclisum = numberFormat(df['gcli'])
        otherFeessum = numberFormat(df['otherFees'])
        monthlyAmountsum = numberFormat(df['monthlyAmount'])
        # df['PNV'] = pd.to_numeric(df['PNV'].fillna(0), errors='coerce')
        # df['PNV'] = df['PNV'].map('{:,.2f}'.format)
        df.sort_values(by=['loanIndex', 'forreleasingdate'], inplace=True)
        df['num'] = numbers(df.shape[0])
        # astype(df, 'loanId', int)
        astype(df, 'term', int)
        astype(df, 'num', int)
        # df[['PNV', 'mlv', 'interest', 'handlingFee', 'dst', 'notarial', 'gcli', 'otherFees', 'monthlyAmount']] = pd.options.display.float_format = '{:,.2f}'.format
        df = df[['num', 'channelName', 'partnerCode', 'outletCode', 'productCode', 'sa', 'loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "mlv", "interest",
                 "handlingFee", "dst", "notarial", "gcli", "otherFees", "term", "actualRate", "monthlyAmount", 'applicationDate', 'approvalDate', 'forreleasingdate', 'fdd',
                 'promoName']]

 # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 33)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        # df50.loc['Total'] = pd.Series(df50['PNV', 'mlv', 'interest', 'handlingFee', 'dst', 'notarial', 'gcli', 'otherFees', 'monthlyAmount'].sum(), index=['PNV', 'mlv', 'interest', 'handlingFee', 'dst', 'notarial', 'gcli', 'otherFees', 'monthlyAmount'])
        # df50.loc['Total'] = df50.PNV.apply(lambda x: "{:,}".format(x))
        df50['num'] = df50['num'].map('{:.0f}'.format)
        df50['term'] = df50['term'].map('{:.0f}'.format)
        # df50['loanId'] = df50['loanId'].map('{:.0f}'.format)
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = 'SUB'
        df50.loc['Total', 'channelName'] = 'TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'term'] = ''

        df50['PNV'] = dfNumberFormat(df50['PNV'])
        df50['mlv'] = dfNumberFormat(df50['mlv'])
        df50['interest'] = dfNumberFormat(df50['interest'])
        df50['handlingFee'] = dfNumberFormat(df50['handlingFee'])
        df50['dst'] = dfNumberFormat(df50['dst'])
        df50['notarial'] = dfNumberFormat(df50['notarial'])
        df50['gcli'] = dfNumberFormat(df50['gcli'])
        df50['otherFees'] = dfNumberFormat(df50['otherFees'])
        df50['monthlyAmount'] = dfNumberFormat(df50['monthlyAmount'])


        # df50.loc['Total']['num'] = df50.loc['Total']['num'].replace(np.float, 'SUB TOTAL', regex=True)
        # print('SUBTOTAL', df50.loc[    'Total']['num'])

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape',
        # 'footer-right': '[page] of [topage]'
    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('booking_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                         name=name, df=split_df_to_chunks_of_50, pnvsum=pnvsum, mlvsum=mlvsum, interestsum=interestsum,
                           handlingFeesum=handlingFeesum, dstsum=dstsum, notarialsum=notarialsum, gclisum=gclisum, otherFeessum=otherFeessum,
                           monthlyAmountsum=monthlyAmountsum, options=options, empty_df=df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Booking Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/incentive", methods=['GET'])
def get_incentive():

    output = BytesIO()

    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')
    name = request.args.get('name')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("generateincentiveReportJSON")

    r = requests.post(url, json=payload)
    data_json = r.json()
    headers = ["#", "BOOKING DATE", "APP ID", "CLIENT'S NAME", "REFERRAL TYPE", "SA", "BRANCH", "LOAN TYPE",  "TERM", "MLV", "PNV",
               "MI", "REFERRER"]
    df = pd.DataFrame(data_json['generateincentiveReportJSONResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 12)))
        sumtotalAmount = ''
        sumPNV = ''
        summonthlyAmount = ''
    else:
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanIndex', int)
        df.sort_values(by=['agentName'], inplace=True)
        df['bookingDate'] = pd.to_datetime(df['bookingDate'])
        df['bookingDate'] = df['bookingDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'bookingDate')
        sumtotalAmount = numberFormat(df['totalAmount'])
        sumPNV = numberFormat(df['PNV'])
        summonthlyAmount = numberFormat(df['monthlyAmount'])
        df = df[['num', 'bookingDate', 'loanId', 'newCustomerName', 'refferalType', "SA", "dealerName", "loanType", "term",
             "totalAmount", "PNV", "monthlyAmount", "agentName"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 38)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['num'] = df['num'].map('{:.0f}'.format)
        # df50['loanId'] = df['loanId'].map('{:.0f}'.format)
        df50['term'] = df['term'].map('{:.0f}'.format)
        df50['totalAmount'] = dfNumberFormat(df50['totalAmount'])
        df50['PNV'] = dfNumberFormat(df50['PNV'])
        df50['monthlyAmount'] = dfNumberFormat(df50['monthlyAmount'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'bookingDate'] = ''
        df50.loc['Total', 'newCustomerName'] = 'SUB TOTAL:'

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'

    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('incentive_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, df=split_df_to_chunks_of_50,
                           sumtotalAmount=sumtotalAmount, sumPNV=sumPNV, summonthlyAmount=summonthlyAmount, empty_df=df.empty)
    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Sales Referral Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/mature", methods=['GET'])
def get_mature():

    output = BytesIO()

    dates = request.args.get('date')
    name = request.args.get('name')

    payload = {'date': dates}

    url = serviceUrl.format("maturedLoanReport")

    r = requests.post(url, json=payload)
    data_json = r.json()

    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "TERM", "BUCKET", "MI", "BMLV",
               "LAST DUE DATE", "LAST PAYMENT", "NO. OF UNPAID", "TOTAL PAYMENT", "TOTAL PAST DUE",
               "TOTAL PENALTY TO PAY", "OB", "NO. OF MONTHS FROM MATURITY"]
    df = pd.DataFrame(data_json['maturedLoanReportResult'])
#    print(data_json)
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 16)))
        sumbMLV = ''
        sumtotalPayment = ''
        summonthlydue = ''
        sumoutStandingBalance = ''
    else:
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        astype(df, 'monthlydue', float)
        astype(df, 'totalPastDue', float)
        astype(df, 'outStandingBalance', float)
        astype(df, 'duePenalty', float)
        astype(df, 'loanIndex', int)
        astype(df, 'unpaidMonths', int)
        astype(df, 'term', int)



        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['loanIndex'], inplace=True)
        dfDateFormat(df, 'lastDueDate')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        sumbMLV = numberFormat(df['bMLV'])
        sumtotalPayment = numberFormat(df['totalPayment'])
        summonthlydue = numberFormat(df['monthlydue'])
        sumoutStandingBalance = numberFormat(df['outStandingBalance'])
        df = df[['num', 'loanId', 'loanAccountNo', 'newCustomerName', "mobileno", "term", "bucket", "monthlydue",
                 "bMLV", "lastDueDate", "lastPayment",
                 "unpaidMonths", "totalPayment", "totalPastDue", "duePenalty", "outStandingBalance", "matured"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 50)
    pd.set_option('display.max_columns', None)
    print(df.head(3))
    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50['num'] = df['num'].map('{:.0f}'.format)
        df50['term'] = df['term'].map('{:.0f}'.format)
        df50['unpaidMonths'] = df['unpaidMonths'].map('{:.0f}'.format)
        df50['monthlydue'] = dfNumberFormat(df50['monthlydue'])
        df50['bMLV'] = dfNumberFormat(df50['bMLV'])
        df50['totalPayment'] = dfNumberFormat(df50['totalPayment'])
        df50['totalPastDue'] = dfNumberFormat(df50['totalPastDue'])
        df50['outStandingBalance'] = dfNumberFormat(df50['outStandingBalance'])
        df50.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'unpaidMonths'] = ''
        df50.loc['Total', 'matured'] = ''
    print(df50.head(2))
    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('mature_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, df=split_df_to_chunks_of_50,
                           sumbMLV=sumbMLV, sumtotalPayment=sumtotalPayment, summonthlydue=summonthlydue,
                           sumoutStandingBalance=sumoutStandingBalance, empty_df=df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Matured Loans Summary as of {}.pdf'.format(dates)

    return response

@pdf_api.route("/duetoday", methods=['GET'])
def get_due():

    output = BytesIO()

    dates = request.args.get('date')
    name = request.args.get('name')

    payload = {'date': dates}

    url = serviceUrl.format("dueTodayReport")

    r = requests.post(url, json=payload)
    data_json = r.json()
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "LOAN TYPE", "DUE TODAY TERM",
               "MI", "TOTAL PAST DUE", "UNPAID PENALTY", "MONTHLY DUE", "LAST PAYMENT DATE", "LAST PAYMENT AMOUNT"]
    df = pd.DataFrame(data_json['dueTodayReportResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 12)))
        summonthlyAmmortization = ''
        summonthdue = ''
        sumunpaidPenalty = ''
        sumlastPaymentAmount = ''
    else:
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        astype(df, 'monthlyAmmortization', float)
        astype(df, 'monthdue', float)
        astype(df, 'loanIndex', int)
        astype(df, 'term', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['loanIndex'], inplace=True)
        dfDateFormat(df, 'monthlydue')
        dfDateFormat(df, 'lastPayment')
        df['num'] = numbers(df.shape[0])
        summonthlyAmmortization = numberFormat(df['monthlyAmmortization'])
        summonthdue = numberFormat(df['monthdue'])
        sumunpaidPenalty = numberFormat(df['unpaidPenalty'])
        sumlastPaymentAmount = numberFormat(df['lastPaymentAmount'])
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileno", "loanType", "term", "monthlyAmmortization",
             "monthdue", "unpaidPenalty", "monthlydue", "lastPayment", "lastPaymentAmount"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 36)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['num'] = df['num'].map('{:.0f}'.format)
        # df50['loanId'] = df['loanId'].map('{:.0f}'.format)
        df50['term'] = df['term'].map('{:.0f}'.format)
        df50['monthlyAmmortization'] = dfNumberFormat(df50['monthlyAmmortization'])
        df50['monthdue'] = dfNumberFormat(df50['monthdue'])
        df50['unpaidPenalty'] = dfNumberFormat(df50['unpaidPenalty'])
        df50['lastPaymentAmount'] = dfNumberFormat(df50['lastPaymentAmount'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'term'] = ''

    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
    }
    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('duetoday_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"), range=xldate_header, time=timeNow,
                           name=name, df=split_df_to_chunks_of_50, summonthlyAmmortization=summonthlyAmmortization,
                           summonthdue=summonthdue, sumunpaidPenalty=sumunpaidPenalty, sumlastPaymentAmount=sumlastPaymentAmount, empty_df=df.empty)
    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Due Today Report {}.pdf'.format(dates)

    return response

@pdf_api.route("/monthlyincome", methods=['GET'])
def get_monthly1():

    output = BytesIO()

    dates = request.args.get('date')
    name = request.args.get('name')
    datetime_object = datetime.strptime(dates, '%m/%d/%Y')
    month = datetime_object.strftime("%B")

    payload = {'date': dates}

    url = serviceUrl.format("monthlyIncomeReportJs")

    r = requests.post(url, json=payload)
    data_json = r.json()

    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "PENALTY PAID",
               "INTEREST PAID", "PRINCIPAL PAID", "UNAPPLIED BALANCE", "PAYMENT AMOUNT", "OR DATE", "OR #"]
    df = pd.DataFrame(data_json['monthlyIncomeReportJsResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 10)))
        sumpenaltyPaid = ''
        suminterestPaid = ''
        sumprincipalPaid = ''
        sumunappliedBalance = ''
        sumpaymentAmount = ''
    else:
        df['loanAccountno'] = df['loanAccountno'].map(lambda x: x.lstrip("'"))
        astype(df, 'loanIndex', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['loanIndex', 'orDate'], inplace=True)
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'orDate')
        df = round(df, 2)
        sumpenaltyPaid = numberFormat(df['penaltyPaid'])
        suminterestPaid = numberFormat(df['interestPaid'])
        sumprincipalPaid = numberFormat(df['principalPaid'])
        sumunappliedBalance = numberFormat(df['unappliedBalance'])
        sumpaymentAmount = numberFormat(df['paymentAmount'])
        df = df[['num', 'appId', 'loanAccountno', 'newCustomerName', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
                 'paymentAmount', "orDate", "orNo"]]

    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 36)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['num'] = df['num'].map('{:.0f}'.format)
        # df50['appId'] = df['appId'].map('{:.0f}'.format)
        df50['penaltyPaid'] = dfNumberFormat(df50['penaltyPaid'])
        df50['interestPaid'] = dfNumberFormat(df50['interestPaid'])
        df50['principalPaid'] = dfNumberFormat(df50['principalPaid'])
        df50['unappliedBalance'] = dfNumberFormat(df50['unappliedBalance'])
        df50['paymentAmount'] = dfNumberFormat(df50['paymentAmount'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanAccountno'] = 'SUB TOTAL:'
        df50.loc['Total', 'appId'] = ''
        df50.loc['Total', 'num'] = ''

    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "For the month of {}".format(month)

    # pass list of dataframes to template
    temp = render_template('monthlyincome_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, df=split_df_to_chunks_of_50,
                           sumpenaltyPaid=sumpenaltyPaid, sumpaymentAmount=sumpaymentAmount, sumprincipalPaid=sumprincipalPaid,
                           suminterestPaid=suminterestPaid, sumunappliedBalance=sumunappliedBalance, empty_df=df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Monthly Income {}.pdf'.format(dates)

    return response

@pdf_api.route("/unappliedbalances", methods=['GET'])
def get_uabalances():
    output = BytesIO()

    name = request.args.get('name')
    dates = request.args.get('date')
    payload = {}

    url = serviceUrl.format("accountDueReportJSON")

    r = requests.post(url, json=payload)
    data = r.json()

    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "AMOUNT DUE", "DUE DATE",
               "UNAPPLIED BALANCE"]
    df = pd.DataFrame(data['accountDueReportJSONResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 7)))
        sumamountDue = ''
        sumunappliedBalance = ''
    else:
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanIndex', int)
        df.sort_values(by=['loanIndex'], inplace=True)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['dueDate'] = pd.to_datetime(df['dueDate'])
        df['dueDate'] = df['dueDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'dueDate')
        sumamountDue = numberFormat(df['amountDue'])
        sumunappliedBalance = numberFormat(df['unappliedBalance'])
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]

    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 50)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['num'] = df['num'].map('{:.0f}'.format)
        # df50['loanId'] = df['loanId'].map('{:.0f}'.format)
        df50['amountDue'] = dfNumberFormat(df50['amountDue'])
        df50['unappliedBalance'] = dfNumberFormat(df50['unappliedBalance'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'
        df50.loc['Total', 'num'] = ''

    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('unappliedbalances_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, df=split_df_to_chunks_of_50,
                           sumunappliedBalance=sumunappliedBalance, sumamountDue=sumamountDue, empty_df=df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Unapplied Balance Summary {}.pdf'.format(dates)

    return response

@pdf_api.route("/newmemoreport", methods=['GET'])
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
        creditDf = pd.DataFrame(pd.np.empty((0, 14)))
    else:
        astype(creditDf, 'loanIndex', int)
        creditDf.sort_values(by=['loanIndex'], inplace=True)
        creditDf['loanAccountNo'] = creditDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        # creditDf['date'] = creditDf.date.apply(lambda x: x.split(" ")[0])
        dfDateFormat(creditDf, 'approvedDate')
        dfDateFormat(creditDf, 'date')
        creditDf['num'] = numbers(creditDf.shape[0])
        creditDf = creditDf[["num", "appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                             "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]

    debitDf = pd.DataFrame(data['Debit'])

    if debitDf.empty:
        debitDf = pd.DataFrame(pd.np.empty((0, 14)))
    else:
        astype(debitDf, 'loanIndex', int)
        debitDf.sort_values(by=['loanIndex'], inplace=True)
        debitDf['loanAccountNo'] = debitDf['loanAccountNo'].map(lambda x: x.lstrip("'"))
        dfDateFormat(debitDf, 'approvedDate')
        dfDateFormat(debitDf, 'date')
        debitDf['num'] = numbers(debitDf.shape[0])
        debitDf = debitDf[["num", "appId", "loanAccountNo", "fullName", "subProduct", "memoType", "purpose", "amount",
                           "status", "date", "createdBy", "remark", "approvedDate", "approvedBy", "approvedRemark"]]

# split the dataframe into rows of 50
    split_creditDf_to_chunks_of_50 = split_dataframe_to_chunks(creditDf, 50)
    # add Totals row to each dataframe
    for df1 in split_creditDf_to_chunks_of_50:
        # df1.loc['Total'] = df1.select_dtypes(pd.np.number).sum()
        df1.loc['Total'] = round(df1.select_dtypes(pd.np.number).sum(), 2)
        df1['num'] = df1['num'].map('{:.0f}'.format)
        df1['appId'] = df1['appId'].map('{:.0f}'.format)
        df1.loc['Total'] = df1.loc['Total'].replace(np.nan, '', regex=True)
        df1.loc['Total', 'appId'] = ''
        df1.loc['Total', 'num'] = ''
        df1.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'


    split_debitDf_to_chunks_of_50 = split_dataframe_to_chunks(debitDf, 50)
    # add Totals row to each dataframe
    for df2 in split_debitDf_to_chunks_of_50:
        df2.loc['Total'] = round(df1.select_dtypes(pd.np.number).sum(), 2)
        df2['num'] = df1['num'].map('{:.0f}'.format)
        df2['appId'] = df1['appId'].map('{:.0f}'.format)
        df2.loc['Total'] = df1.loc['Total'].replace(np.nan, '', regex=True)
        df2.loc['Total', 'appId'] = ''
        df2.loc['Total', 'num'] = ''
        df2.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'

    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('memo_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                           name=name, creditDfs=split_creditDf_to_chunks_of_50, debitDfs=split_debitDf_to_chunks_of_50,
                           empty_credit=creditDf.empty, empty_debit=debitDf.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Memo Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/tat", methods=['GET'])
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

    standard_df = pd.read_csv(StringIO(standard))
    returned_df = pd.read_csv(StringIO(returned))

    dfDateFormat(standard_df, 'Application Date')
    dfDateFormat(returned_df, 'Application Date')

    standard_df.insert(0, column='#', value=numbers(standard_df.shape[0]))
    returned_df.insert(0, column='#', value=numbers(returned_df.shape[0]))

    sumMLV = numberFormat(standard_df['MLV'])
    sumPNV = numberFormat(standard_df['PNV'])
    sumPenVer = round(pd.Series(standard_df['Pending - For Verification']).sum(), 2)
    sumVerAdj = round(pd.Series(standard_df['For Verification - For Adjudication']).sum(), 2)
    sumVerCan = round(pd.Series(standard_df['For Verification - For Cancellation']).sum(), 2)
    sumCanCan = round(pd.Series(standard_df['For Cancellation - Cancelled']).sum(), 2)
    sumAdjApr = round(pd.Series(standard_df['For Adjudication - For Approval']).sum(), 2)
    sumAprApr = round(pd.Series(standard_df['For Approval - Approved']).sum(), 2)
    sumAprDis = round(pd.Series(standard_df['For Approval - Disapproved']).sum(), 2)
    sumAprRel = round(pd.Series(standard_df['Approved - For Releasing']).sum(), 2)
    sumRelRel = round(pd.Series(standard_df['For Releasing - Released']).sum(), 2)

# split the dataframe into rows of 50
    split_standard_df_to_chunks_of_50 = split_dataframe_to_chunks(standard_df, 25)
    # add Totals row to each dataframe
    for df50 in split_standard_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['#'] = standard_df['#'].map('{:.0f}'.format)
        # df50['App ID'] = standard_df['App ID'].map('{:.0f}'.format)
        df50['MLV'] = dfNumberFormat(df50['MLV'])
        df50['PNV'] = dfNumberFormat(df50['PNV'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'First Name'] = 'SUB TOTAL:'
        df50.loc['Total', '#'] = ''
        df50.loc['Total', 'App ID'] = ''

    split_returned_df_to_chunks_of_50 = split_dataframe_to_chunks(returned_df, 25)
    # add Totals row to each dataframe
    for df50 in split_returned_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50['#'] = returned_df['#'].map('{:.0f}'.format)
        # df50['App ID'] = returned_df['App ID'].map('{:.0f}'.format)
        df50['MLV'] = dfNumberFormat(df50['MLV'])
        df50['PNV'] = dfNumberFormat(df50['PNV'])
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'First Name'] = 'SUB TOTAL:'
        df50.loc['Total', '#'] = ''
        df50.loc['Total', 'App ID'] = ''

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'

    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('tat_template.html', standardHeaders=standardHeaders, returnedHeaders=returnedHeaders, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, standardDF=split_standard_df_to_chunks_of_50, sumMLV=sumMLV, sumPNV=sumPNV,
                           returnedDF=split_returned_df_to_chunks_of_50, sumPenVer=sumPenVer, sumVerAdj=sumVerAdj,
                           sumVerCan=sumVerCan, sumCanCan=sumCanCan, sumAdjApr=sumAdjApr, sumAprApr=sumAprApr, sumAprDis=sumAprDis,
                           sumAprRel=sumAprRel, sumRelRel=sumRelRel, empty_standard=standard_df.empty, empty_returned=returned_df.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=TAT {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/customerLedger", methods=['GET'])
def get_customerLedger():

    output = BytesIO()

    loanId = request.args.get('loanId')
    userId = request.args.get('userId')
    dates = request.args.get('date')
    name = request.args.get('name')

    payload = {'loanId': loanId, 'userId': userId, 'date': dates}

    ledgerById = "customerLedger/ledgerByLoanId?loanId={}".format(loanId)
    url = bluemixUrl.format(ledgerById)
    url2 = serviceUrl.format("getCustomerLedger")

    r = requests.post(url2, json=payload)
    data_json = r.json()
    ledgerData = requests.get(url).json()

    dfLedger = pd.DataFrame(ledgerData['data']['transactions'])
    dfCustomerLedger = pd.DataFrame(data_json['getCustomerLedgerResult'])

    headers = ["#", "DATE", "TERM", "TRANSACTION TYPE", "PAYMENT TYPE", "REF NO", "CHECK #", "PENALTY INCUR", "PRINCIPAL",
               "INTEREST", "PENALTY PAID", "ADVANCES", "TOTAL", "DUE", "OB", "PAYMENT DATE", "OR NO", "OR DATE"]

    headersBorrower = ["APPLICATION ID", "LOAN ACCOUNT NO.", "BORROWER'S NAME", "COLLECTOR", "CONTACT NO.", "ADDRESS"]
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


    prinPaid = dfSummary['principalPaid'] - dfSummary['creditPrincipal']
    intPaid = dfSummary['interestPaid'] - dfSummary['creditInterest']
    penPaid = dfSummary['penaltyPaid'] - dfSummary['creditPenalty']

    prinTotal = dfSummary['principal'] - dfSummary['debitPrincipal']
    intTotal = dfSummary['interest'] - dfSummary['debitInterest']
    penTotal = dfSummary['penalty'] - dfSummary['debitPenalty']

    gTotal = round(dfSummary['principal'] + dfSummary['interest'] + dfSummary['penalty'], 2)
    gTotalPaid = round(dfSummary['principalPaid'] + dfSummary['interestPaid'] + dfSummary['penaltyPaid'], 2)
    gTotalAdj = round(dfSummary['principalAdj'] + dfSummary['interestAdj'] + dfSummary['penaltyAdj'], 2)
    gTotalBilled = round(dfSummary['principalBilled'] + dfSummary['interestBilled'] + dfSummary['penaltyBilled'], 2)
    gTotalAmtDue = round(dfSummary['principalAmtDue'] + dfSummary['interestAmtDue'] + dfSummary['penaltyAmtDue'], 2)
    gTotalBal = round((dfSummary['principal'] + dfSummary['interest'] + dfSummary['penalty']) - (dfSummary['principalPaid'] + dfSummary['interestPaid'] + dfSummary['penaltyPaid']) + (dfSummary['principalAdj'] + dfSummary['interestAdj'] + dfSummary['penaltyAdj']), 2)

    total = round(dfSummary['principal'] + dfSummary['interest'], 2)
    totalPaid = round(dfSummary['principalPaid'] + dfSummary['interestPaid'], 2)
    totalAdj = round(dfSummary['principalAdj'] + dfSummary['interestAdj'], 2)
    totalBilled = round(dfSummary['principalAdj'] + dfSummary['interestAdj'], 2)
    totalAmtDue = round(dfSummary['principalAmtDue'] + dfSummary['interestAmtDue'], 2)
    totalBal = round((dfSummary['principal'] + dfSummary['interest']) - (dfSummary['principalPaid'] + dfSummary['interestPaid']) + (dfSummary['principalAdj'] + dfSummary['interestAdj']), 2)

    principalBal = round((dfSummary['principal'] - dfSummary['debitPrincipal']) - (dfSummary['principalPaid'] - dfSummary['creditPrincipal']) + (dfSummary['debitPrincipal'] + (dfSummary['creditPrincipal'] * -1)), 2)
    interestBal = round((dfSummary['interest'] - dfSummary['debitInterest']) - (dfSummary['interestPaid'] - dfSummary['creditInterest']) + (dfSummary['debitInterest'] + (dfSummary['creditInterest'] * -1)), 2)
    penaltyBal = round((dfSummary['penalty'] - dfSummary['debitPenalty']) - (dfSummary['penaltyPaid'] - dfSummary['creditPenalty']) + (dfSummary['debitPenalty'] + (dfSummary['creditPenalty'] * -1)), 2)

    list1 = [len(i) for i in headers]

    if dfLedger.empty:
        dfLedger = pd.DataFrame(pd.np.empty((0, 12)))
    else:
        dfLedger['orDate'] = dfLedger['orDate'].loc[dfLedger['orDate'].str.contains("/")]
        dfLedger['paymentDate'] = dfLedger['paymentDate'].loc[dfLedger['paymentDate'].str.contains("/")]
        conditions = [(dfLedger['paymentDate'] == '-')]
        dfLedger['paymentDate'] = np.select(conditions, [dfLedger['paymentDate']], default="")
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
        dfLedger['orDate'] = dfLedger['orDate'].replace(np.nan, '-', regex=True)
        dfLedger = round(dfLedger, 2)
        dfLedger = dfLedger[["num", "date", "term", "type", "paymentType", "refNo", "checkNo", "penaltyIncur", "principalPaid", "interestPaid", "penaltyPaid",
             "advances", "totalRow", "amountDue", "ob", "paymentDate", "orNo", "orDate"]]
    dfCustomerLedger = round(dfCustomerLedger, 2)

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(dfLedger, 50)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # print('borrowerDetails', dfCustomerLedger['borrower'])
    # pass list of dataframes to template
    temp = render_template('customerledger_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                           name=name, df=split_df_to_chunks_of_50, headersBorrower=headersBorrower,
                           borrowerDetails=dfCustomerLedger['borrower'], headersAccStat=headersAccStat, headersOB=headersOB,
                           headersLoan=headersLoan, headersCollateral=headersCollateral, headersLoanSummary=headersLoanSummary,
                           dataAcctStat=dfCustomerLedger['acctStat'], dataLoan=dfCustomerLedger['loan'],
                           dataCollateral=dfCustomerLedger['collateral'], dataAcctSum=dfCustomerLedger['accountSummary'],
                           adjPrincipal=adjPrincipal, adjInterest=adjInterest, adjPenalty=adjPenalty,
                           prinPaid=prinPaid, intPaid=intPaid, penPaid=penPaid,
                           prinTotal=prinTotal, intTotal=intTotal, penTotal=penTotal,
                           total=total, totalPaid=totalPaid, totalAdj=totalAdj, totalBilled=totalBilled, totalAmtDue=totalAmtDue,
                           totalBal=totalBal, gTotal=gTotal, gTotalPaid=gTotalPaid, gTotalAdj=gTotalAdj, gTotalBilled=gTotalBilled, gTotalAmtDue=gTotalAmtDue,
                           gTotalBal=gTotalBal, principalBal=principalBal, interestBal=interestBal, penaltyBal=penaltyBal, empty_df=dfLedger.empty)

    config = _get_pdfkit_config()

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename="Customer Ledger {}.pdf'.format(loanId)

    return response

def split_dataframe_to_chunks(df, n):
    df_len = len(df)
    count = 0
    dfs = []

    while True:
        if count > df_len-1:
            break

        start = count
        count += n
        #print("%s : %s" % (start, count))
        dfs.append(df.iloc[start : count])
    return dfs
			
