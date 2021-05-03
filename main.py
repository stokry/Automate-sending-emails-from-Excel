import xlrd
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

path = "clients.xlsx"
openFile = xlrd.open_workbook(path)
sheet = openFile.sheet_by_name('clients')


mail_list = []
amount = []
name = []
for k in range(sheet.nrows-1):
    client = sheet.cell_value(k+1,0)
    email = sheet.cell_value(k+1,1)
    paid = sheet.cell_value(k+1,3)
    count_amount = sheet.cell_value(k+1,4)
    if paid == 'No':
        mail_list.append(email) 
        amount.append(count_amount)
        name.append(client)


email = 'some@gmail.com' 
password = 'pass' 
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, password)

for mail_to in mail_list:
    send_to_email = mail_to
    find_des = mail_list.index(send_to_email) 
    clientName = name[find_des] 
    subject = f'{clientName} you have a new email'
    message = f'Dear {clientName}, \n' \
              f'we inform you that you owe ${amount[find_des]}. \n'\
              '\n' \
              'Best Regards' 

    msg = MIMEMultipart()
    msg['From '] = send_to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    text = msg.as_string()
    print(f'Sending email to {clientName}... ') 
    server.sendmail(email, send_to_email, text)

server.quit()
print('Process is finished!')
time.sleep(10) 