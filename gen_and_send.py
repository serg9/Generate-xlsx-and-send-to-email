import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email import encoders
from email.mime.base import MIMEBase
import xlsxwriter
import time


workbook = xlsxwriter.Workbook('Mapper.xlsx')
worksheet = workbook.add_worksheet()

data = (
    ['Code', 'Name', 'Source', 'Dictionary'],
    ['ОС00021694', 'Name 1', 'FIN', 'product'],
    ['ОС00021694', 'Name 2', 'FIN', 'product'],
    ['ОС00021694', 'Name 3', 'FIN', 'product'],
    ['ОС00021694', 'Name 4', 'FIN', 'product'],
)
row = 0
col = 0

for Code, Name, Source, Dictionary in (data):
    worksheet.write(row, 0, Code)
    worksheet.write(row, 1, Name)
    worksheet.write(row, 2, Source)
    worksheet.write(row, 3, Dictionary)
    row += 1
workbook.close()

emailfrom = "" #Add your
emailto = "" #Add your
fileToSend = "Test.xlsx"
username = "" #Add your
password = "your pass" #Add your

msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg["Subject"] = "Title message"
msg.preamble = "preamble"

ctype, encoding = mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    ctype = "application/octet-stream"

maintype, subtype = ctype.split("/", 1)
count_elements = len(data)
has_data = """ We have """ + str(count_elements)
withou_data = " We don't have elements in table "
result_check_data = ""
if len(data) > 0:
    fp = open(fileToSend, "rb")
    attachment = MIMEBase(maintype, subtype)
    attachment.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
    msg.attach(attachment)
    result_check_data = has_data
else:
    result_check_data = withou_data
body = """
<div dir="ltr" id="divtagdefaultwrapper" style="color:black;font-size:12pt;font-family:Calibri,Helvetica,sans-serif;">
    <div style="color:black;">
        <div dir="ltr" id="divRplyFwdMsg">
            <font color="black" face="Calibri,sans-serif" style="font-size:11pt;"><b>From:</b> reports@my-domain.com &lt;reports@my-domain.com&gt;<br>
                <b>Sent:</b> """+ str(time.strftime("%a, %d %b %Y %H:%M:%S")) +"""<br>
                <b>To:</b> """ + str(emailto) + """<br>
                <b>Cc:</b> Your row;<br>
                <b>Subject:</b>""" + str(result_check_data) + """Good day to you! ;)
            </font>
        </div>  
    </div>
</div>
"""
msg.attach(MIMEText(body, "html","UTF-8")) 

server = smtplib.SMTP("smtp.gmail.com:587")
server.starttls()
server.login(username, password)
server.sendmail(emailfrom, emailto, msg.as_string())
print('sendmail')
server.quit()
