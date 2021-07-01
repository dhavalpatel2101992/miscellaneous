import os
import openpyxl
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

datevalue=datetime.datetime.today().strftime('%Y%m%d')
timestamp = datetime.datetime.today().strftime('%Y-%m-%d_%H:%M:%S')
print('*****************************************')
print('Script Run DateTime: ',timestamp)


# Change current working dir
wrk_dir=r'E:\Control-M\tabprd\FPM\Python_Scripts\\'
os.chdir(wrk_dir)

# reading excel
wb = openpyxl.load_workbook('Database_Details.xlsx')
sheet = wb["DatabaseDetails"]
filepath = sheet['B6'].value
filename = sheet['B7'].value
filename = filepath+'\\'+filename+'_'+datevalue+'.xlsx'
from_id = sheet['B8'].value
fail_mailid = sheet['B9'].value
user_tomailid = sheet['B10'].value
user_ccmailid = sheet['B11'].value
bcc_id=fail_mailid
subject=sheet['B12'].value
mailbody=sheet['B13'].value

s = smtplib.SMTP(host='smtphost.qualcomm.com', port=25)
s.starttls()

def sendErrormail(errormsg):
    msg = MIMEMultipart()
    msg['From'] = from_id
    msg['To'] = fail_mailid
    msg['Subject'] = 'Action Required: Delivery Failed - FCPA New Company Report V3'
    message = timestamp+''': Error Occurred while generating FCPA New Company Report V3 in Python Code.
    Error Details:'''+str(errormsg)
    msg.attach(MIMEText(message, 'plain'))

    try:
        s.send_message(msg)
        print("Developers have been notified about Failure...")
    except:
        print("Mail Server Error - Error Occurred Email not sent to Internal team")
    del msg
def sendnotificationmail():
    msg = MIMEMultipart()
    msg['From'] = from_id
    msg['To'] = user_tomailid
    msg['Cc'] = user_ccmailid
    msg['Bcc'] = bcc_id
    msg['Subject'] = subject
    message = mailbody
    msg.attach(MIMEText(message, 'plain'))

    try:
        s.send_message(msg)
        print("Users have been notified Successfully...")
    except Exception as e:
        print("Mail Server Error - Email not sent to users")
        print('Error Details:', e)
        print('About to send Failure Notification to Developers...')
        sendErrormail(e)
    del msg

try:

    if os.path.isfile(filename):
        print('About to send Report Notification to Users...')
        sendnotificationmail()
    else:
        print('Error Occurred while generating FCPA New Company Report V3 in Python Code')
        print('Error Details: File Not Found')
        print('About to send Failure Notification to Developers...')
        sendErrormail('File Not Found')

except Exception as e:
    print('Error Occurred while generating FCPA New Company Report V3 in Python Code')
    print('Error Details:',e)
    print('About to send Failure Notification to Developers...')
    sendErrormail(e)

s.quit()
