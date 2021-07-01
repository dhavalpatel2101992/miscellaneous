from datetime import datetime,timedelta
import re
import os
import time
import openpyxl
from zeep import Client
from requests import Session
from zeep.transports import Transport
import xml.etree.ElementTree as ET
from zeep.wsse.username import UsernameToken
import pandas as pd
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font,Alignment,PatternFill
from cryptography.fernet import Fernet
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import shutil

timestamp = datetime.today().strftime('%Y-%m-%d_%H:%M:%S')
print('*' * 60)
print('Script Start DateTime: ', timestamp)

# Change current working dir
wrk_dir=r'E:\Control-M\tabprd\FPM\OBIEE_Python_Expense_Report_Details\\'
os.chdir(wrk_dir)

def calFromTo():
    """
    The function checks current date and returns From date & To date
    """
    curr_date=datetime.now()
    prev_month_enddt = curr_date.replace(day=1) - timedelta(days=1)
    curr_quarter_startdt = datetime(curr_date.year, 3 * pd.Timestamp(curr_date).quarter - 2, 1)
    prev_quarter_startdt = datetime(curr_date.year, 3 * pd.Timestamp(prev_month_enddt).quarter - 2, 1)
    
    if curr_date-curr_quarter_startdt > timedelta(days=15):
        fromdt=curr_quarter_startdt.strftime("%Y-%m-%d")
        if int(curr_date.strftime("%d"))>15:
            todt=curr_date.strftime("%Y-%m-15")
        else:
            todt=prev_month_enddt.strftime("%Y-%m-%d")
    else:
        fromdt=prev_quarter_startdt.strftime("%Y-%m-%d")
        todt=prev_month_enddt.strftime("%Y-%m-%d")
    return fromdt,todt
def sendReportGenNotification():
    s = smtplib.SMTP(host='smtphost.qualcomm.com', port=25)
    s.starttls()
    msg = MIMEMultipart()
    msg['From'] = from_id
    msg['To'] = to_id
    msg['Cc'] = cc_id
    msg['Bcc'] = bcc_id
    msg['Subject'] = 'New version of {0} is available'.format(reportname)
    html_str ="""
    <html>
    <body>
    <style> a {font-size: 16px}; table.main {width: 740px;} td {font-family: 'Microsoft Sans Serif';}</style>
    <table class="main" cellspacing="0" cellpadding="10">
    <tbody>
    <tr>
    <td style="font-weight: bold;font-size: 18px;padding:20px;background-color:#bdcfff;">
    Notification : """+ """{0} ({1}-{2})""".format(reportname,fromdt,todt)+"""
    </td>
    </tr>
    <tr>
    <td style="font-size: 14px;border:2px solid #bdcfff; padding: 30px 30px 30px 40px;">
    New version of """+reportname+""" is available to you at: 
    <br>"""+filepath+"""
    <br><br>
    This report contains ERs which have been submitted between """+fromdt+""" and """+todt+""".
    </td>
    </tr>
    <tr>
    <td style="font-size: 14px;padding:12px;background-color:#bdcfff;">
    &nbsp; For questions or help, please send an email to <a href="mailto:bi.help@qualcomm.com">bi.help@qualcomm.com</a>
    </td>
    </tr>
    </tbody>
    </table>
    </body>
    </html>"""
    msg.attach(MIMEText(html_str, 'html'))
    try:
        s.send_message(msg)
        print("Developments & Users have been notified about Report Generation")
    except:
        print("Mail Server Error - Report Generation Notification not sent to Developments")
    del msg
    s.quit()
    return None
def sendErrorNotification(errormsg):
    s = smtplib.SMTP(host='smtphost.qualcomm.com', port=25)
    s.starttls()
    msg = MIMEMultipart()
    msg['From'] = from_id
    msg['To'] = bcc_id
    msg['Subject'] = 'Action Required: Report Generation Failed - '+reportname
    message = timestamp+': Error Occurred while generating {0}. Error Details: {1}'.format(reportname,errormsg)
    msg.attach(MIMEText(message, 'plain'))
    try:
        s.send_message(msg)
        print("Development team has been notified about Failure")
    except:
        print("Mail Server Error - Failure Notification not sent to Development team")
        print('*' * 60)
        sys.exit(0)
    del msg
    s.quit()
    print('*' * 60)
    sys.exit(0)

# reading excel
wb = openpyxl.load_workbook('Bursting_Details.xlsx')

# mail-server details
sheet = wb["Server&LoginDetails"]
from_id = sheet['C6'].value
bcc_id = sheet['C7'].value
to_id = sheet['C9'].value
cc_id = sheet['C10'].value
filepath = sheet['C11'].value
reportname = 'Expense_Report_Details'
# From Date & To Date
fromdt, todt= calFromTo()
print('From:', fromdt, 'To:', todt)
filename = reportname + '_' + fromdt + ' to ' + todt + '.xlsx'
if os.path.isfile(filepath + '\\' + filename):
    print('About to send Report Generation Notification...')
    sendReportGenNotification()
else:
    print('Error Occurred while generating {0}'.format(reportname))
    print('Error Details: File Not Found')
    print('About to send Failure Notification to Developers...')
    sendErrorNotification('File Not Found')
timestamp = datetime.today().strftime('%Y-%m-%d_%H:%M:%S')
print('Script Completion DateTime: ', timestamp)
print('*' * 60)