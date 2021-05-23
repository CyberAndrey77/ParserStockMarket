import mimetypes
import os.path
import smtplib
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_message(email, message, filePath):
    fromAddr = "fine.andrey231@yandex.ru"
    toAddr = email
    password = "uvkulakuxrvzjumr"

    msg = MIMEMultipart()
    msg['From'] = fromAddr
    msg['To'] = toAddr
    msg['Subject'] = 'TestTask'

    messageBody = message
    msg.attach(MIMEText(messageBody, 'plain'))
    attach_file(msg, filePath)

    server = smtplib.SMTP_SSL("smtp.yandex.ru", 465)
    server.login(fromAddr, password)
    server.send_message(msg)
    server.quit()


def attach_file(msg, filePath):
    fileName = os.path.basename(filePath)
    ctype, encoding = mimetypes.guess_type(fileName)
    if ctype is None or encoding is None:
        ctype = 'application/octet-stream'
    mainType, subType = ctype.split('/', 1)
    if mainType == 'text':
        with open(filePath) as fp:
            file = MIMEText(fp.read(), _subtype=subType)
            fp.close()
    elif mainType == 'image':
        with open(filePath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subType)
            fp.close()
    elif mainType == 'audio':
        with open(filePath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subType)
            fp.close()
    else:
        with open(filePath, 'rb') as fp:
            file = MIMEBase(mainType, subType)
            file.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(file)
    file.add_header('Content-Disposition', 'attachment', filename=fileName)
    msg.attach(file)

def main():
    send_message('lapardin.andrey@mail.ru', 'Test', 'USD_EUR.xlsx')

if __name__ == '__main__':
    main()