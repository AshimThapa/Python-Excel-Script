from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path as op

def sendEmail(paths,file_path):
    with open(file_path+'Email/EmailList.txt','r') as email_file:
        email_list=email_file.readlines()

    msg=MIMEMultipart()
    sender='sender@email.com'
    subject='ALERT !!! Main Page missing file'
    body='Hi\n'+'There has been multiple trip files during this month.\n'+'Please refer to the attached file\n\n\n'+'Regards\n'+'Auto SYS.'
    msg['From']=sender
    msg['To']=','.join(email_list)
    msg['Subject']=subject
    msg.attach(MIMEText(body,'plain'))



    for path in paths:
        part=MIMEBase('application','octet-stream')
        with open(path,'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment;filename="{}"'.format(op.basename(path)))
        msg.attach(part)

    s=SMTP('your.smtp.com')
    s.sendmail(sender,email_list,msg.as_string())
    s.quit()
    return
