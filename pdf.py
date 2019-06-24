from flask import Flask, Blueprint, request, jsonify, send_file, render_template, make_response
import json
import requests
import pandas as pd
import numpy as np
import openpyxl
import flask_excel as excel
from io import BytesIO, StringIO
import os

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

serviceUrl = "https://api360.zennerslab.com/Service1.svc/{}" #rfc-service-live
lambdaUrl = "https://ia-lambda-live.mybluemix.net/{}" #lambda-pivotal-live
bluemixUrl = "https://rfc360.mybluemix.net/{}" #rfc-bluemix-live

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

@pdf_api.route("/")
def pdfList():
    return "list of pdfs"


@pdf_api.route("/collectionPDF", methods=['GET'])
def collection_pdf():
    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("collection")
    print(url)
    r = requests.post(url, json=payload)
    data = r.json()
   
    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "AMT DUE", "FDD", "PNV", "MLV", "MI", "TERM", "PEN", "INT",
               "PRIN", "UNPAID MOS", "PAID MOS"]
    df = pd.DataFrame(data['collectionResult'])
    list1 = [len(i) for i in headers]
    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 21)))
    else:
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
        sumamountDue = round(pd.Series(df['amountDue']).sum(), 2)
        sumpnv = round(pd.Series(df['pnv']).sum(), 2)
        summlv = round(pd.Series(df['mlv']).sum(), 2)
        summi = round(pd.Series(df['mi']).sum(), 2)
        sumsumOfPenalty = round(pd.Series(df['sumOfPenalty']).sum(), 2)
        sumtotalInterest = round(pd.Series(df['totalInterest']).sum(), 2)
        sumtotalPrincipal = round(pd.Series(df['totalPrincipal']).sum(), 2)
        sumhf = round(pd.Series(df['hf']).sum(), 2)
        sumdst = round(pd.Series(df['dst']).sum(), 2)
        sumnotarial = round(pd.Series(df['notarial']).sum(), 2)
        sumgcli = round(pd.Series(df['gcli']).sum(), 2)
        sumoutstandingBalance = round(pd.Series(df['outstandingBalance']).sum(), 2)
        sumtotalPayment = round(pd.Series(df['totalPayment']).sum(), 2)
        df = df[["num","loanId", "loanAccountNo", "name", "amountDue", "fdd", "pnv", "mlv", "mi", "term",
                 "sumOfPenalty", "totalInterest", "totalPrincipal", "unapaidMonths", "paidMonths", "hf", "dst", "notarial", "gcli", "outstandingBalance", "status",
                 "totalPayment"]]

    # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 37)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
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
                           name=name, df=split_df_to_chunks_of_50, sumamountDue=sumamountDue,
                           sumpnv=sumpnv, summlv=summlv, summi=summi, sumsumOfPenalty=sumsumOfPenalty, sumtotalInterest=sumtotalInterest,
                           sumtotalPrincipal=sumtotalPrincipal, sumhf=sumhf, sumdst=sumdst, sumnotarial=sumnotarial, sumgcli=sumgcli,
                           sumoutstandingBalance=sumoutstandingBalance, sumtotalPayment=sumtotalPayment)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options,configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Collection Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response


@pdf_api.route("/dccrPDF", methods=['GET'])
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
        sumamount = round(pd.Series(df['amount']).sum(), 2)
        sumcash = round(pd.Series(df['cash']).sum(), 2)
        sumpaymentCheck = round(pd.Series(df['paymentCheck']).sum(), 2)
        sumpaidPrincipal = round(pd.Series(df['paidPrincipal']).sum(), 2)
        sumpaidInterest = round(pd.Series(df['paidInterest']).sum(), 2)
        sumadvances = round(pd.Series(df['advances']).sum(), 2)
        sumpaidPenalty = round(pd.Series(df['paidPenalty']).sum(), 2)
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
    split_df = split_dataframe_to_chunks(df, 38)
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
        df1.loc['Total'] = df1.loc['Total'].replace(np.nan, '', regex=True)
        df1.loc['Total', 'num'] = ''
        df1.loc['Total', 'transType'] = 'SUB TOTAL:'

    for df2 in split_dfCash:
        df2.loc['Total'] = round(df1.select_dtypes(pd.np.number).sum(), 2)
        df2.loc['Total'] = df2.loc['Total'].replace(np.nan, '', regex=True)
        df2.loc['Total', 'dfCashnum'] = 'SUB TOTAL:'
    for df3 in split_dfEcpay:
        df3.loc['Total'] = round(df3.select_dtypes(pd.np.number).sum(), 2)
        df3.loc['Total'] = df3.loc['Total'].replace(np.nan, '', regex=True)
        df3.loc['Total', 'dfEcpaynum'] = 'SUB TOTAL:'
    for df4 in split_dfBC:
        df4.loc['Total'] = round(df4.select_dtypes(pd.np.number).sum(), 2)
        df4.loc['Total'] = df4.loc['Total'].replace(np.nan, '', regex=True)
        df4.loc['Total', 'dfBCnum'] = 'SUB TOTAL:'
    for df5 in split_dfBank:
        df5.loc['Total'] = round(df5.select_dtypes(pd.np.number).sum(), 2)
        df5.loc['Total'] = df5.loc['Total'].replace(np.nan, '', regex=True)
        df5.loc['Total', 'dfBanknum'] = 'SUB TOTAL:'
    for df6 in split_dfCheck:
        df6.loc['Total'] = round(df6.select_dtypes(pd.np.number).sum(), 2)
        df6.loc['Total'] = df6.loc['Total'].replace(np.nan, '', regex=True)
        df6.loc['Total', 'dfChecknum'] = 'SUB TOTAL:'
    for df7 in split_dfGPRS:
        df7.loc['Total'] = round(df7.select_dtypes(pd.np.number).sum(), 2)
        df7.loc['Total'] = df7.loc['Total'].replace(np.nan, '', regex=True)
        df7.loc['Total', 'dfGPRSnum'] = 'SUB TOTAL:'
    for df8 in split_df2:
        df8.loc['Total'] = round(df8.select_dtypes(pd.np.number).sum(), 2)
        df8.loc['Total'] = df8.loc['Total'].replace(np.nan, '', regex=True)
        df8.loc['Total', 'num'] = 'SUB TOTAL:'

    print(split_df)

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('dccr_template.html', headers=headers1, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                           name=name, df=split_df , split_dfCash=split_dfCash, split_dfEcpay=split_dfEcpay, split_dfBC=split_dfBC,
                           split_dfBank=split_dfBank, split_dfCheck=split_dfCheck, split_dfGPRS=split_dfGPRS, split_df2=split_df2,
                           sumamount=sumamount, sumcash=sumcash, sumpaymentCheck=sumpaymentCheck, sumpaidPrincipal=sumpaidPrincipal,
                           sumpaidInterest=sumpaidInterest, sumadvances=sumadvances, sumpaidPenalty=sumpaidPenalty)
    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=DCCR {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/agingPDF", methods=['GET'])
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
    else:
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
        agingp1DF["newCustomerName"] = agingp1DF['lastName'] + ', ' + agingp1DF['firstName'] + ' ' + agingp1DF[
            'middleName'] + ' ' + agingp1DF['suffix']
        agingp1DF['ob'] = agingp1DF['notDue'] + agingp1DF['monthDue']
        agingp1DF = round(agingp1DF, 2)
        summonthlyInstallment = round(pd.Series(agingp1DF['monthlyInstallment']).sum(), 2)
        sumob = round(pd.Series(agingp1DF['ob']).sum(), 2)
        sumrunningMLV = round(pd.Series(agingp1DF['runningMLV']).sum(), 2)
        sumtoday = round(pd.Series(agingp1DF['today']).sum(), 2)
        sum1 = round(pd.Series(agingp1DF['1-30']).sum(), 2)
        sum31 = round(pd.Series(agingp1DF['31-60']).sum(), 2)
        sum61 = round(pd.Series(agingp1DF['61-90']).sum(), 2)
        sum91 = round(pd.Series(agingp1DF['91-120']).sum(), 2)
        sum121 = round(pd.Series(agingp1DF['121-150']).sum(), 2)
        sum151 = round(pd.Series(agingp1DF['151-180']).sum(), 2)
        sum181 = round(pd.Series(agingp1DF['181-360']).sum(), 2)
        sum360 = round(pd.Series(agingp1DF['360 & over']).sum(), 2)
        sumtotal = round(pd.Series(agingp1DF['total']).sum(), 2)
        sumduePrincipal = round(pd.Series(agingp1DF['duePrincipal']).sum(), 2)
        dueInterest = round(pd.Series(agingp1DF['dueInterest']).sum(), 2)
        sumduePenalty = round(pd.Series(agingp1DF['duePenalty']).sum(), 2)
        sumamountSum = round(pd.Series(agingp1DF['amountSum']).sum(), 2)
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
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'channelName'] = 'SUB TOTAL:'
        df50.loc['Total', 'appId'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'expiredTerm'] = ''

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('aging_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"), range=xldate_header,
                           time=timeNow,name=name, df=split_df_to_chunks_of_50, summonthlyInstallment=summonthlyInstallment,
                           sumob=sumob, sumrunningMLV=sumrunningMLV, sumtoday=sumtoday, sum1=sum1, sum31=sum31, sum61=sum61,
                           sum91=sum91, sum121=sum121, sum151=sum151, sum181=sum181, sum360=sum360, sumtotal=sumtotal,
                           sumduePrincipal=sumduePrincipal, dueInterest=dueInterest, sumduePenalty=sumduePenalty, sumamountSum=sumamountSum)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Aging Report as of {}.pdf'.format(dates)

    return response

@pdf_api.route("/bookingPDF", methods=['GET'])
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
        astype(df, 'loanId', int)
        astype(df, 'term', int)
        pnvsum = round(pd.Series(df['PNV']).sum(), 2)
        mlvsum = round(pd.Series(df['mlv']).sum(), 2)
        interestsum = round(pd.Series(df['interest']).sum(), 2)
        handlingFeesum = round(pd.Series(df['handlingFee']).sum(), 2)
        dstsum = round(pd.Series(df['dst']).sum(), 2)
        notarialsum = round(pd.Series(df['notarial']).sum(), 2)
        gclisum = round(pd.Series(df['gcli']).sum(), 2)
        otherFeessum = round(pd.Series(df['otherFees']).sum(), 2)
        monthlyAmountsum = round(pd.Series(df['monthlyAmount']).sum(), 2)
        df.sort_values(by=['loanId', 'forreleasingdate'], inplace=True)
        df['num'] = numbers(df.shape[0])
        df = df[['num', 'channelName', 'partnerCode', 'outletCode', 'productCode', 'sa', 'loanId', 'loanAccountNo', 'customerName', "subProduct", "PNV", "mlv", "interest",
                 "handlingFee", "dst", "notarial", "gcli", "otherFees", "term", "actualRate", "monthlyAmount", 'applicationDate', 'approvalDate', 'forreleasingdate', 'fdd',
                 'promoName']]

 # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 38)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        countdf50 =df50.shape[0]
        # df50.loc['Total'] = pd.Series(df50['PNV', 'mlv', 'interest', 'handlingFee', 'dst', 'notarial', 'gcli', 'otherFees', 'monthlyAmount'].sum(), index=['PNV', 'mlv', 'interest', 'handlingFee', 'dst', 'notarial', 'gcli', 'otherFees', 'monthlyAmount'])
        # df50.loc['Total'] = df50.PNV.apply(lambda x: "{:,}".format(x))
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = 'SUB'
        df50.loc['Total', 'channelName'] = 'TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'term'] = ''
        print('df50', df50)
        # df50.loc['Total']['num'] = df50.loc['Total']['num'].replace(np.float, 'SUB TOTAL', regex=True)
        # print('SUBTOTAL', df50.loc['Total']['num'])

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
                           monthlyAmountsum=monthlyAmountsum, options=options, countdf50=countdf50)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Booking Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/incentivePDF", methods=['GET'])
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
    else:
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['agentName'], inplace=True)
        df['bookingDate'] = pd.to_datetime(df['bookingDate'])
        df['bookingDate'] = df['bookingDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'bookingDate')
        sumtotalAmount = round(pd.Series(df['totalAmount']).sum(), 2)
        sumPNV = round(pd.Series(df['PNV']).sum(), 2)
        summonthlyAmount = round(pd.Series(df['monthlyAmount']).sum(), 2)
        df = df[['num', 'bookingDate', 'loanId', 'newCustomerName', 'refferalType', "SA", "dealerName", "loanType", "term",
             "totalAmount", "PNV", "monthlyAmount", "agentName"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 40)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'bookingDate'] = 'SUB TOTAL:'

    print(split_df_to_chunks_of_50)

    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'

    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('incentive_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow, name=name, df=split_df_to_chunks_of_50,
                           sumtotalAmount=sumtotalAmount, sumPNV=sumPNV, summonthlyAmount=summonthlyAmount)
    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Sales Referral Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/maturePDF", methods=['GET'])
def get_mature():

    output = BytesIO()

    dates = request.args.get('date')
    name = request.args.get('name')

    payload = {'date': dates}

    url = serviceUrl.format("maturedLoanReport")

    r = requests.post(url, json=payload)
    data_json = r.json()

    headers = ["#", "APP ID", "LOAN ACCT. #", "CLIENT'S NAME", "MOBILE #", "TERM", "BMLV", "LAST DUE DATE",
               "LAST PAYMENT", "NO. OF UNPAID", "TOTAL PAYMENT", "TOTAL PAST DUE", "OB",
               "NO. OF MONTHS"]
    df = pd.DataFrame(data_json['maturedLoanReportResult'])
    list1 = [len(i) for i in headers]

    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 13)))
    else:
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
        sumbMLV = round(pd.Series(df['bMLV']).sum(), 2)
        sumtotalPayment = round(pd.Series(df['totalPayment']).sum(), 2)
        summonthlydue = round(pd.Series(df['monthlydue']).sum(), 2)
        sumoutStandingBalance = round(pd.Series(df['outStandingBalance']).sum(), 2)
        df = df[['num', 'loanId', 'loanAccountNo', 'newCustomerName', "mobileno", "term", "bMLV", "lastDueDate", "lastPayment",
                 "unpaidMonths", "totalPayment", "monthlydue", "outStandingBalance", "matured"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 50)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'term'] = ''
        df50.loc['Total', 'unpaidMonths'] = ''
        df50.loc['Total', 'matured'] = ''

    print(split_df_to_chunks_of_50)

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
                           sumoutStandingBalance=sumoutStandingBalance)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Matured Loans Report as of {}.pdf'.format(dates)

    return response

@pdf_api.route("/duetodayPDF", methods=['GET'])
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

    print(df)
    if df.empty:
        df = pd.DataFrame(pd.np.empty((0, 12)))
    else:
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
        summonthlyAmmortization = round(pd.Series(df['monthlyAmmortization']).sum(), 2)
        summonthdue = round(pd.Series(df['monthdue']).sum(), 2)
        sumunpaidPenalty = round(pd.Series(df['unpaidPenalty']).sum(), 2)
        sumlastPaymentAmount = round(pd.Series(df['lastPaymentAmount']).sum(), 2)
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileno", "loanType", "term", "monthlyAmmortization",
             "monthdue", "unpaidPenalty", "monthlydue", "lastPayment", "lastPaymentAmount"]]

# split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 35)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanAccountNo'] = 'SUB TOTAL:'
        df50.loc['Total', 'loanId'] = ''
        df50.loc['Total', 'num'] = ''
        df50.loc['Total', 'term'] = ''

    print('split_df_to_chunks_of_50', split_df_to_chunks_of_50)

    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
    }
    xldate_header = "As of {}".format(startDateFormat(dates))

    # pass list of dataframes to template
    temp = render_template('duetoday_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"), range=xldate_header, time=timeNow,
                           name=name, df=split_df_to_chunks_of_50, summonthlyAmmortization=summonthlyAmmortization,
                           summonthdue=summonthdue, sumunpaidPenalty=sumunpaidPenalty, sumlastPaymentAmount=sumlastPaymentAmount)
    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Due Today Report {}.pdf'.format(dates)

    return response

@pdf_api.route("/monthlyincomePDF", methods=['GET'])
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
    else:
        df['loanAccountno'] = df['loanAccountno'].map(lambda x: x.lstrip("'"))
        astype(df, 'appId', int)
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        df.sort_values(by=['appId', 'orDate'], inplace=True)
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'orDate')
        df = round(df, 2)
        sumpenaltyPaid = round(pd.Series(df['penaltyPaid']).sum(), 2)
        suminterestPaid = round(pd.Series(df['interestPaid']).sum(), 2)
        sumprincipalPaid = round(pd.Series(df['principalPaid']).sum(), 2)
        sumunappliedBalance = round(pd.Series(df['unappliedBalance']).sum(), 2)
        sumpaymentAmount = round(pd.Series(df['paymentAmount']).sum(), 2)
        df = df[['num', 'appId', 'loanAccountno', 'newCustomerName', "penaltyPaid", "interestPaid", "principalPaid", "unappliedBalance",
                 'paymentAmount', "orDate", "orNo"]]

    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 38)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
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
                           suminterestPaid=suminterestPaid, sumunappliedBalance=sumunappliedBalance)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Monthly Income {}.pdf'.format(dates)

    return response

@pdf_api.route("/unappliedbalancesPDF", methods=['GET'])
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
    else:
        df["newCustomerName"] = df['lastName'] + ', ' + df['firstName'] + ' ' + df['middleName'] + ' ' + df['suffix']
        astype(df, 'loanId', int)
        df.sort_values(by=['loanId'], inplace=True)
        df['loanAccountNo'] = df['loanAccountNo'].map(lambda x: x.lstrip("'"))
        df['dueDate'] = pd.to_datetime(df['dueDate'])
        df['dueDate'] = df['dueDate'].map(lambda x: x.strftime('%m/%d/%Y') if pd.notnull(x) else '')
        df['num'] = numbers(df.shape[0])
        dfDateFormat(df, 'dueDate')
        sumamountDue = round(pd.Series(df['amountDue']).sum(), 2)
        sumunappliedBalance = round(pd.Series(df['unappliedBalance']).sum(), 2)
        df = df[["num", "loanId", "loanAccountNo", "newCustomerName", "mobileNo", "amountDue", "dueDate", "unappliedBalance"]]

    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 50)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'loanId'] = 'SUB TOTAL:'
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
                           sumunappliedBalance=sumunappliedBalance, sumamountDue=sumamountDue)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Unapplied Balance {}.pdf'.format(dates)

    return response

@pdf_api.route("/memoreportPDF", methods=['GET'])
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
        astype(creditDf, 'appId', int)
        creditDf.sort_values(by=['appId'], inplace=True)
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
        astype(debitDf, 'appId', int)
        debitDf.sort_values(by=['appId'], inplace=True)
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
        df1.loc['Total'] = df1.select_dtypes(pd.np.number).sum()

    split_debitDf_to_chunks_of_50 = split_dataframe_to_chunks(debitDf, 50)
    # add Totals row to each dataframe
    for df2 in split_debitDf_to_chunks_of_50:
        df2.loc['Total'] = df2.select_dtypes(pd.np.number).sum()

    print('CREDIT', split_creditDf_to_chunks_of_50)
    print('DEBIT', split_debitDf_to_chunks_of_50)
    options = {
        'page-size': 'Legal',
        'orientation': 'Landscape'

    }

    xldate_header = "Period: {}-{}".format(startDateFormat(dateStart), endDateFormat(dateEnd))

    # pass list of dataframes to template
    temp = render_template('memo_template.html', headers=headers, date=date.today().strftime("%m/%d/%y"),
                           range=xldate_header, time=timeNow,
                           name=name, creditDf=split_creditDf_to_chunks_of_50, debitDf=split_debitDf_to_chunks_of_50)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=Memo Report {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/tatPDF", methods=['GET'])
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

    sumMLV = round(pd.Series(standard_df['MLV']).sum(), 2)
    sumPNV = round(pd.Series(standard_df['PNV']).sum(), 2)
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
    split_standard_df_to_chunks_of_50 = split_dataframe_to_chunks(standard_df, 38)
    # add Totals row to each dataframe
    for df50 in split_standard_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
        df50.loc['Total'] = df50.loc['Total'].replace(np.nan, '', regex=True)
        df50.loc['Total', 'First Name'] = 'SUB TOTAL:'
        df50.loc['Total', '#'] = ''
        df50.loc['Total', 'App ID'] = ''

    split_returned_df_to_chunks_of_50 = split_dataframe_to_chunks(returned_df, 38)
    # add Totals row to each dataframe
    for df50 in split_returned_df_to_chunks_of_50:
        df50.loc['Total'] = round(df50.select_dtypes(pd.np.number).sum(), 2)
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
                           sumAprRel=sumAprRel, sumRelRel=sumRelRel)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

    pdf = pdfkit.from_string(temp, False, options=options, configuration=config)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=TAT {}-{}.pdf'.format(dateStart, dateEnd)

    return response

@pdf_api.route("/customerledgerPDF", methods=['GET'])
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

    print(split_df_to_chunks_of_50)

    options = {
        # 'page-size': 'Legal',
        'orientation': 'Landscape'
        #
    }

    xldate_header = "As of {}".format(startDateFormat(dates))

    # print('borrowerDetails', dfCustomerLedger['borrower'])
    print('borrowerDetails', dfCustomerLedger)
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
                           gTotalBal=gTotalBal, principalBal=principalBal, interestBal=interestBal, penaltyBal=penaltyBal)

    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

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