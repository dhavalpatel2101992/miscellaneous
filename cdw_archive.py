import shutil
import os
import smtplib
from email.mime.text import MIMEText

# network drive details
SrcFiles=r'/prj/infa_corpfin_prd/edw9_r12/SrcFiles'
User_SrcFiles=SrcFiles+'/CDW/User_SrcFiles'
archive=SrcFiles+'/CDW/FILE_LIST/archive'

# mail-server details:
host = 'smtphost.qualcomm.com'
port = 25
sender = 'IMPORT COMPLIANCE <no.reply@qualcomm.com>'
recipients = ['INFORMATICA_FPM_CDM_USERS@qti.qualcomm.com']
subject = 'Customs Import processed Files'


# source files
csvfiles=[]
for fname in os.listdir(User_SrcFiles):
    if os.path.isdir(User_SrcFiles + '/' + fname):
        pass
    else:
        if fname[-4:].lower()=='.csv' and fname[0:2] != '~$':
            csvfiles.append(fname)
print('{0} csv files found:'.format(len(csvfiles)))
print(csvfiles)

# processed files
processedcsvfile=open(SrcFiles+'/CDW/F_ACT_REP_IMPORT_DTL_FILE_NAMES.txt', "r")
processedfiles = processedcsvfile.read().replace('"','').splitlines()

# archive folder checking
if not os.path.exists(archive):
    os.makedirs(archive)

# files to be archived
files_to_move = [file for file in processedfiles if file in csvfiles]

# archive process
for file in files_to_move:
    try:
        shutil.move(User_SrcFiles + "/"+ file, archive +"/"+ file)
        shutil.move(User_SrcFiles + "/"+ file.replace('.csv','.xlsx'), archive +"/"+ file.replace('.csv','.xlsx'))
    except Exception as e:
        print(str(e))

# mail body
if len(files_to_move) > 0:
    msgstring = "Following Files have been Processed and Archived Successfully: "+"\n" + "\n".join(files_to_move)
else:
    msgstring = "No Files have been Processed or Archived."

# smtp server initialization
s = smtplib.SMTP(host=host, port=port)
s.starttls()

# smtp msg building
msg = MIMEText(msgstring)
msg['From'] = sender
msg['To'] = ";".join(recipients)
msg['Bcc']='dhavpate@qti.qualcomm.com'
msg['Subject'] = subject

# mail sending
try:
    s.sendmail(sender,recipients, msg.as_string())
    print(msgstring)
    print("Email has been sent to recipients")
except Exception as e:
    print(str(e))
    print("Error - Email not sent")
del msg
s.quit()