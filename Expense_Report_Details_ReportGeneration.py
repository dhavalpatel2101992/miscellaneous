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
    # curr_date = datetime(2021, 1, 2, 16, 6, 27, 553884)
    prev_month_enddt = curr_date.replace(day=1) - timedelta(days=1)
    curr_quarter_startdt = datetime(curr_date.year, 3 * pd.Timestamp(curr_date).quarter - 2, 1)
    year_startdt = datetime(curr_date.year, 1, 1)
    if curr_date - year_startdt > timedelta(days=15):
        prev_quarter_startdt = datetime(curr_date.year, 3 * pd.Timestamp(prev_month_enddt).quarter - 2, 1)
    else:
        prev_quarter_startdt = datetime(curr_date.year - 1, 3 * pd.Timestamp(prev_month_enddt).quarter - 2, 1)
    
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

def createSummary(df):
    cols=['ER Number','Business Unit','Division Name','Employee Department', 'Manager Name',
       'FA Name', 'Employee Name', 'Employee Title',
       'Employee Country', 'Employee Number', 'Submit Date','Total Authorized Amount', 'Total Authorized Currency','Exported Date']
    return df.fillna('-').groupby(cols, as_index=False).agg({'Approved Amount (USD)':'sum'})


def writeFile(df,fromdt,todt,filepath,reportname):
    """
    Remove older files if exists
    """
    for fname in os.listdir(filepath):
        if os.path.isdir(filepath + '/' + fname):
            pass
        else:
            if fname[-5:].lower()=='.xlsx' and fname[0:2] != '~$' and fromdt in fname:
                os.remove(filepath+'\\'+fname)
                print('Deleted {0}'.format(fname))
    """
    The function writes dataframe in formatted excel.
    """
    filename = reportname + '_' + fromdt + ' to ' + todt + '.xlsx'
    xls_writer = pd.ExcelWriter(filename, engine='xlsxwriter',date_format='YYYY-MM-DD',datetime_format='YYYY-MM-DD',options={'strings_to_numbers': True})
    df.to_excel(xls_writer, sheet_name='Detail Report', float_format="%.2f", index=False)
    workbook = xls_writer.book
    worksheet = xls_writer.sheets['Detail Report']
    # Add a header format.
    header_format = workbook.add_format({'bold': True, 'bg_color': '#BFD2E2', 'border': 1})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = df[value].astype(str).str.len().max()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(value)) + 1
        # set the column length
        worksheet.set_column(col_num, col_num, column_len)
    #xls_writer.save()

    if (not df.empty):
        summarydf=createSummary(df)
    else:
        summarydf=pd.DataFrame(columns=['No Data Available'])
    summarydf.to_excel(xls_writer, sheet_name='Summary Report', float_format="%.2f", index=False)
    worksheet = xls_writer.sheets['Summary Report']
    # Add a header format.
    header_format = workbook.add_format({'bold': True, 'bg_color': '#BFD2E2', 'border': 1})
    for col_num, value in enumerate(summarydf.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = summarydf[value].astype(str).str.len().max()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(value)) + 1
        # set the column length
        worksheet.set_column(col_num, col_num, column_len)
    xls_writer.save()

    shutil.move(filename, filepath + '\\' + filename)
    if os.path.isfile(filepath + '\\' + filename):
        print('About to send Report Generation Notification...')
        sendReportGenNotification()
    return None
def downloadReportDf(wsdl,username,password,reportpath,fromdt,todt):
    """
    The function connects to the OBIEE and download report's data, returnes Pandas DataFrame
    """
    print('Downloading Started')
    executionOptions={"async" : True, "maxRowsPerPage" : 100, "refresh" : True, "presentationInfo" : True}
    # Initializing SOAP client, start a session, and make a connection to XmlViewService binding
    session = Session()
    session.verify = True
    transport = Transport(session=session)
    client = Client(wsdl=wsdl, wsse=UsernameToken(username, password), transport=transport)
    sessionid = client.service.logon(username, password)
    xmlservice = client.bind('XmlViewService')


    # Retrieveing data schema and column headings
    max_retries = 30
    while max_retries > 0:
        schema = xmlservice.executeXMLQuery(report=reportpath, outputFormat="SAWRowsetSchema", executionOptions=executionOptions, sessionID=sessionid)
        if schema.rowset == None:
            max_retries -= 1
            continue
        else:
            time.sleep(10)
            break

    if schema.rowset == None:
        client.service.logoff(sessionID=sessionid)
        raise Exception("SAWRowsetSchemaError")

    columnHeading = re.findall(r'columnHeading="(.*?)"', schema.rowset)
    dataset_dict = {}

    for head in columnHeading:
        dataset_dict[head] = []
    dataset_dict.pop('Line Item Id')  #hidden columns  
    
    
    P_DTRANGE_FROM="DATE '{}'".format(fromdt) #"DATE '2020-06-01'"
    P_DTRANGE_TO="DATE '{}'".format(todt)   #"DATE '2020-06-30'"
    reportParams={"filterExpressions":"","variables":[{"name":"P_DTRANGE_FROM","value":P_DTRANGE_FROM},{"name":"P_DTRANGE_TO","value":P_DTRANGE_TO}]}
    # Making a query and parsing first datarows
    queryresult = xmlservice.executeXMLQuery(report=reportpath, outputFormat="SAWRowsetData",
                                                executionOptions=executionOptions,reportParams=reportParams, sessionID=sessionid)    
    queryid = queryresult.queryID

    if queryresult.rowset == None:
        client.service.logoff(sessionID=sessionid)
        raise Exception("SAWRowsetDataError")

    ETobject = ET.fromstring(queryresult.rowset)
    namespacerows = ETobject.findall('{urn:schemas-microsoft-com:xml-analysis:rowset}Row')

    for row in namespacerows:
        for key in dataset_dict.keys():
            try:
                cellvalue=row.find('{urn:schemas-microsoft-com:xml-analysis:rowset}Column' + 
                                                  str(list(dataset_dict.keys()).index(key))).text
            except:
                cellvalue=None
            dataset_dict[key].append(cellvalue)


    # Determine if additional fetching is needed and if yes - parsing additional rows   
    queryfetch = queryresult.finished

    while (not queryfetch):
        queryfetch = xmlservice.fetchNext(queryID=queryid, sessionID=sessionid)
        ETobject = ET.fromstring(queryfetch.rowset)
        namespacerows = ETobject.findall('{urn:schemas-microsoft-com:xml-analysis:rowset}Row')

        for row in namespacerows:
            for key in dataset_dict.keys():
                try:
                    cellvalue=row.find('{urn:schemas-microsoft-com:xml-analysis:rowset}Column' + 
                                                  str(list(dataset_dict.keys()).index(key))).text
                except:
                    cellvalue=None
                dataset_dict[key].append(cellvalue)

        queryfetch = queryfetch.finished

    # By some reason OBIEE doesn't make the last fetching, it will fix it
    queryfetch = False

    while (not queryfetch):
        queryfetch = xmlservice.fetchNext(queryID=queryid, sessionID=sessionid)
        ETobject = ET.fromstring(queryfetch.rowset)
        namespacerows = ETobject.findall('{urn:schemas-microsoft-com:xml-analysis:rowset}Row')

        for row in namespacerows:
            for key in dataset_dict.keys():
                try:
                    cellvalue=row.find('{urn:schemas-microsoft-com:xml-analysis:rowset}Column' + 
                                                  str(list(dataset_dict.keys()).index(key))).text
                except:
                    cellvalue=None
                dataset_dict[key].append(cellvalue)
        queryfetch = True

    reportdf=pd.DataFrame(dataset_dict)
    # date columns
    for col in reportdf.columns:
        if 'date' in col.lower():
            reportdf[col] = pd.to_datetime(reportdf[col], errors='coerce')
    reportdf['Approved Amount (USD)']=reportdf['Approved Amount (USD)'].astype('float')
    print('Downloading Completed')
    return reportdf
def sendReportGenNotification():
    s = smtplib.SMTP(host='smtphost.qualcomm.com', port=25)
    s.starttls()
    msg = MIMEMultipart()
    msg['From'] = from_id
    msg['To'] = bcc_id
    msg['Subject'] = '{0} Generated Successfully'.format(reportname,fromdt,todt)
    message = '''{0}:{1}  -  {2} has been placed successfully at {3}'''.format(reportname,fromdt,todt,filepath)
    msg.attach(MIMEText(message, 'plain'))
    try:
        s.send_message(msg)
        print("Developments have been notified about Report Generation")
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
wsdl = sheet['C1'].value
username = sheet['C2'].value
password = sheet['C3'].value
from_id = sheet['C6'].value
bcc_id = sheet['C7'].value
to_id = sheet['C9'].value
cc_id = sheet['C10'].value
filepath = sheet['C11'].value
# Parameters
reportname = 'Expense_Report_Details'
reportpath = '/shared/Extensity/_portal/Extensity T&E Reporting/Analysis Reports/Expense Report Details Report/Expense Report Details'
key = b'icX0uAcHik9UCEgzTY3jP2_KhbEfXGAZucdSX3sbQMQ='
cipher_suite = Fernet(key)
password = cipher_suite.decrypt(bytes(password, 'utf-8')).decode('utf-8')
print('Parameter Reading Completed')
# From Date & To Date
fromdt, todt= calFromTo()
print('From:', fromdt, 'To:', todt)
# Download Report Dataframe
try:
    reportdf = downloadReportDf(wsdl, username, password, reportpath, fromdt, todt)
except Exception as e:
    print('Error Occurred while downloading Report')
    print('Error Details:', e)
    print('About to send Report Generation Failure Notification')
    sendErrorNotification(str(e))

try:
    writeFile(reportdf, fromdt, todt,filepath, reportname)
    print('Report Saved')
except Exception as e:
    print('Error Occurred while saving Report')
    print('Error Details:', e)
    print('About to send Report Generation Failure Notification')
    sendErrorNotification(str(e))
timestamp = datetime.today().strftime('%Y-%m-%d_%H:%M:%S')
print('Script Completion DateTime: ', timestamp)
print('*' * 60)