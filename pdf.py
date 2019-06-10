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


pdf_api = Blueprint('pdf_api', __name__)

serviceUrl = "https://api360.zennerslab.com/Service1.svc/{}" #rfc-service-live

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

@pdf_api.route("/")
def pdfList():
    return "list of pdfs"


@pdf_api.route("/collectionreport", methods=['GET'])
def collectionreport():
    #same stuff without excel specific
    name = request.args.get('name')
    dateStart = request.args.get('startDate')
    dateEnd = request.args.get('endDate')

    payload = {'startDate': dateStart, 'endDate': dateEnd}

    url = serviceUrl.format("collection")
    print(url)
    r = requests.post(url, json=payload)
    data = r.json()
   
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

    # split the dataframe into rows of 50
    split_df_to_chunks_of_50 = split_dataframe_to_chunks(df, 50)

    # add Totals row to each dataframe
    for df50 in split_df_to_chunks_of_50:
        df50.loc['Total'] = df50.select_dtypes(pd.np.number).sum()
       
    print(split_df_to_chunks_of_50)
    
    options = {
        'orientation': 'Landscape'
    }

    # pass list of dataframes to template
    temp = render_template('report_template.html', headers=headers, date=date.today().strftime('%d, %b %Y'), df=split_df_to_chunks_of_50)
    
    pdf = pdfkit.from_string(temp, False, options=options)
    # respond with PDF
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename="test.pdf'

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