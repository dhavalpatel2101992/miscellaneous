import cx_Oracle
from pandas import DataFrame as df
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import HTML
import os
import openpyxl
import datetime
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
dbuserid = sheet['B1'].value
dbpassword = sheet['B2'].value
dbhost = sheet['B3'].value
dbport = sheet['B4'].value
dbSID = sheet['B5'].value
dbtns=cx_Oracle.makedsn(dbhost,dbport,dbSID)

connection = cx_Oracle.connect(dbuserid,dbpassword,dbtns)
cursor = connection.cursor()
querystring = '''
SELECT PDS.GROUP_NAME,
       (SELECT LISTAGG (
                   A.EMPLOYEE_USERID || '@' || A.EMPLOYEE_EMAIL_DOMAIN, ';')
                   WITHIN GROUP (ORDER BY 1)
          FROM EDW_STAGING.S_COMP_PDS_GROUP_MEMBER A
         WHERE A.GROUP_NAME = PDS.GROUP_NAME AND A.C_ADMINISTRATOR = 'true')
           AS GROUP_ADMIN,
       EMP.FIRST_NAME || ' ' || EMP.LAST_NAME
           MEMBER,
       MNGR.FIRST_NAME || ' ' || MNGR.LAST_NAME
           MANAGER,
       TRUNC (SYSDATE) - TRUNC (PDS.EDW_CREATE_DT)
           DAYS,
       CASE
           WHEN TRUNC (SYSDATE) - TRUNC (PDS.EDW_CREATE_DT) = 1
           THEN
               'Please ensure they have been removed within the next 24 hours.'
           WHEN TRUNC (SYSDATE) - TRUNC (PDS.EDW_CREATE_DT) = 2
           THEN
               'Please take action and remove them ASAP.'
           WHEN TRUNC (SYSDATE) - TRUNC (PDS.EDW_CREATE_DT) = 3
           THEN
               'This is the last call to remove them from the group.'
           WHEN TRUNC (SYSDATE) - TRUNC (PDS.EDW_CREATE_DT) > 3
           THEN
              'Remove them immediately and have justification prepared.'
       END
           EMAIL_BODY
  FROM D_EMPLOYEE  EMP,
       D_EMPLOYEE  MNGR,
       (SELECT *
          FROM EDW_STAGING.S_COMP_PDS_GROUP_MEMBER_INT
         WHERE     TRUNC (SYSDATE) - TRUNC (EDW_CREATE_DT) >= 1
               AND GROUP_NAME IN ('informatica-corp.edw-administrator',
                                  'TABLEAU_FPM_SITE_ADMIN',
                                  'OBIEE_EDW_ADMIN',
                                  'TABLEAU_FPM_REVENUE_REPORTING_ADMIN',
                                  'COG_TM1_QTL_IT_ADMIN')
               AND C_MEMBER = 'true'
               AND EMPLOYEE_TYPE_DESC <> 'Account') PDS
 WHERE     EMP.SUPERVISOR_EMPLOYEE_NUM = MNGR.NK_EMPLOYEE_NUM
       AND EMP.NK_EMPLOYEE_NUM = PDS.EMPLOYEE_NUM
 '''
cursor.execute(querystring)
dataframe = df(cursor.fetchall())
cursor.close()
connection.close()

if(not dataframe.empty):
    s = smtplib.SMTP(host='smtphost.qualcomm.com', port=25)
    s.starttls()
    newdataframe = dataframe.groupby(0)
    for GROUP_NAME, GROUP_NAME_DF in newdataframe:
        print('Processing ', GROUP_NAME, ' ..........')
        msg = MIMEMultipart()
        msg['From'] = 'edwadmin@qualcomm.com'
        msg['To'] = GROUP_NAME_DF.values[0][1].replace('standardsml2','bwigley')
        msg['Bcc']= 'dhavpate@qti.qualcomm.com'
        msg['Subject'] = "Action Required: Remove Member(s) from " + GROUP_NAME
        html_str1 = """
                    <html>
                    <body><p style="font-size:50px;"><center style="font-size:20px;">Remove following member(s) from """ + \
                    '<a href="https://lists.qualcomm.com/ListManager?match=eq&field=default&query='+GROUP_NAME+'">' \
                    + GROUP_NAME + '</a>' +"""<left></p>"""
        html_str3 = """</body>
                    </html>
                    """
        htmlcode = ""

        t = HTML.Table(header_row=['Member', 'Manager', 'No of Days', 'Action Required'])
        for row in range(0, GROUP_NAME_DF.shape[0]):
            t.rows.append([GROUP_NAME_DF.values[row][2], GROUP_NAME_DF.values[row][3], str(GROUP_NAME_DF.values[row][4]), GROUP_NAME_DF.values[row][5]])

        htmlcode = str(t)
        html_str2 = htmlcode
        html_str = html_str1 + html_str2 + html_str3

        msg.attach(MIMEText(html_str, 'html'))
        try:
            s.send_message(msg)
            print("Email successfully sent for ", GROUP_NAME)
        except:
            print("Error - Email not sent for ", GROUP_NAME)
        del msg
    s.quit()
    print('Code Executed Successfully')
else:
    print('No Members found in Admin Groups')
    print('Code Executed Successfully')


