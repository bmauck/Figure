import os
import json
import pandas as pd
import numpy as np
import datetime as dt
import blpapi, pdblp
from pandas.tseries.offsets import BDay

pd.set_option('display.max_columns', 500)
cwd = os.getcwd()
tday = dt.datetime.today().date()
yest = (tday - BDay(1)).date()
yest_str = yest.isoformat().replace('-', '')
tday_str = tday.isoformat().replace('-', '')

cxn = pdblp.BCon(debug=False, port=8194, timeout=5000)
cxn.start()

terms = [5, 7, 10, 15, 20]

px_grid = cwd + '\\slr\\grids\\pricing_fixed_{}yr-20200317.csv'

p4w_close = tday - BDay(20)
p4w_close = p4w_close.date()
p4w_close_str = p4w_close.isoformat().replace('-', '')

p12w_close = tday - BDay(60)
p12w_close = p12w_close.date()
p12w_close_str = p12w_close.isoformat().replace('-', '')

sectys = ['USSW2 Curncy', 'USSW3 Curncy', 'USSW4 Curncy', 'USSW5 Curncy', 'USSW10 Curncy'
            ,'IBOXUMAE CBBT Curncy', 'IBOXHYSE CBBT Curncy', 'SPX Index', 'CCMP Index']

quotes = {}
for sec in sectys:
    print(sec)
    quotes[sec] = cxn.bdh('{}'.format(sec), 'PX_LAST', p4w_close_str, p4w_close_str).values[0][0]

px_file = '\\slr\\pricing_slr-{}.xlsx'.format(tday_str)
writer = pd.ExcelWriter(os.getcwd() + px_file, engine='xlsxwriter')
count = 0
for term in terms:
    df = pd.read_csv(px_grid.format(term), index_col=0)
        
    if term == 5:
        px = (quotes['USSW2 Curncy'] + (quotes['IBOXUMAE CBBT Curncy'] / 100)) / 100
    elif term == 7:
        px = (quotes['USSW2 Curncy'] + (quotes['IBOXUMAE CBBT Curncy'] / 100)) / 100
    elif term == 10:
        px = (quotes['USSW3 Curncy'] + (quotes['IBOXUMAE CBBT Curncy'] / 100)) / 100
    elif term == 15:
        px = (quotes['USSW3 Curncy'] + (quotes['IBOXUMAE CBBT Curncy'] / 100)) / 100
    elif term == 20:
        px = (quotes['USSW4 Curncy'] + (quotes['IBOXUMAE CBBT Curncy'] / 100)) / 100
    else:
        print('This loan term cannot be handled')

    df = df.add(px)
    df1 = df - 0.0164
    
    df.to_excel(writer, sheet_name='px_fixed', startrow=count+1)
    df1.to_excel(writer, sheet_name='px_variable', startrow=count+1)
    
    wb = writer.book
    ws = writer.sheets['px_fixed']
    ws1 = writer.sheets['px_variable']
#     print(df)

    ws.write('A{}'.format(count+1), '{}Y Fixed'.format(term))
    ws.write('A40', 'Assumptions: ' + json.dumps(quotes))
    
    ws1.write('A{}'.format(count+1), '{}Y Variable'.format(term))
    ws1.write('A40', 'Assumptions: ' + json.dumps(quotes))
    
    count += 7
    
wb = writer.book
ws = writer.sheets['px_fixed']
ws1 = writer.sheets['px_variable']


teaser_rate = 0.0349
libor = 0.0164

ws.write('D6', teaser_rate)
ws1.write('D6', teaser_rate - libor)

pct_format = wb.add_format({'num_format': '0.000%'})
ws.set_column('B:D', None, pct_format)
ws1.set_column('B:D', None, pct_format)

wb.close()
writer.save()

import smtplib
import ssl
import email

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sender_email = 'bmauck@figure.com'
receiver_email = 'bmauck@figure.com'
pw = input('Password:')

import yagmail

to = 'bmauck@figure.com'
body = """
Student Loan Refi pricing attached. 


Assumptions: 
{}
""".format(quotes)
 
filename = 'Z:/Shared/Capital Markets/Pricing/slr/pricing_slr-{}.xlsx'.format(tday_str)

yag = yagmail.SMTP('bmauck@figure.com')
yag.send(to=to, 
        subject='Figure SLR Pricing {}'.format(tday_str), 
        contents=body,
        attachments=filename)