
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
import xlsxwriter


workbook = xlsxwriter.Workbook('Test.xlsx')
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
msg["Subject"] = "help I cannot send an attachment to save my life"
msg.preamble = "help I cannot send an attachment to save my life"

ctype, encoding = mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    ctype = "application/octet-stream"

maintype, subtype = ctype.split("/", 1)

fp = open(fileToSend, "rb")
attachment = MIMEBase(maintype, subtype)
attachment.set_payload(fp.read())
fp.close()
encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
msg.attach(attachment)

server = smtplib.SMTP("smtp.gmail.com:587")
server.starttls()
server.login(username, password)
server.sendmail(emailfrom, emailto, msg.as_string())
server.quit()
