# encoding=utf8
import smtplib,email,email.encoders,email.mime.text,email.mime.base
import os
import zipfile
import glob
import csv
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from xlsxwriter.workbook import Workbook
import psycopg2



## delete only if file exists ##
if os.path.exists(''):
    os.remove('')
else:
    print("Sorry, I can not remove " )


#Email Connection Details

smtpserver = 'your smtp fqdn'
to = ['to email_address']
fromAddr = 'from email address'
subject = "Here you go "



query = """ Enter your query """



conn1 = psycopg2.connect("dbname='database_name' user='user-name' host='db_ip_address' password='db_password'")
cur = conn1.cursor()

outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(query)

################################temparary file to store the query output #######################
with open('filename_query_out', 'w') as f:
    cur.copy_expert(outputquery, f)
################################temparary file to store the query output #######################

################################convert your temparary file to xlsx format #######################
for csvfile in glob.glob(os.path.join('.', 'filename*')):
    workbook = Workbook(csvfile + '.xlsx')
    worksheet = workbook.add_worksheet('sheet_name enter here')
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)

    workbook.close()
################################convert your temparary file to xlsx format #######################


######################Remove your temparay file that was used to store query input######################
os.remove('filename_query_out');
######################Remove your temparay file that was used to store query input######################

######################Zip the xlsx file ######################
zip = zipfile.ZipFile('lamda_1.zip', 'a')
zip.write('filename_query_out.xlsx')
zip.close()
######################Zip the xlsx file ######################


######################HTML EMAIL Content######################
# create html email
html = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
html +='"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
html +='<body style="font-size:12px;font-family:Verdana"><p>Here you go</p>'
html += "</body></html>"
emailMsg = email.MIMEMultipart.MIMEMultipart('alternative')
emailMsg['Subject'] = subject
emailMsg['From'] = fromAddr
emailMsg['To'] = ', '.join(to)
emailMsg.attach(email.mime.text.MIMEText(html,'html'))
######################HTML EMAIL Content######################


##########################Attach the zip content ##############################################
fileMsg = email.mime.base.MIMEBase('application','zip')
fileMsg.set_payload(file('lamda_1.zip').read())
email.encoders.encode_base64(fileMsg)
fileMsg.add_header('Content-Disposition','attachment;filename=lamda_1.zip')
emailMsg.attach(fileMsg)
##########################Attach the zip content ##############################################

# send email
server = smtplib.SMTP(smtpserver)
server.sendmail(fromAddr,to,emailMsg.as_string())
server.quit()

