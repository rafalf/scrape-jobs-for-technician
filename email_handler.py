# http://naelshiab.com/tutorial-send-email-python/

import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
import os


def _send_email(recipient, logger, address_urls, cfg):

    email_recipient = recipient

    if cfg['gmail_method'] == 'less-secure':
        gmail_user = cfg['gmail_user']
        gmail_pwd = cfg['gmail_pwd']
        email_sender = cfg['gmail_user']
        email_subject = cfg['email_subject']
        email_content = cfg['email_content']

        for i, urls in enumerate(address_urls):
            email_content += "\n\n--------------------------------------------------"
            email_content += "\nJob: "
            email_content += str(i + 1)
            email_content += "\n\nAddress:\n"
            email_content += urls[0]
            email_content += "\n"
            email_content += "\nFDH address:\n"
            email_content += urls[1]
            email_content += "--------------------------------------------------"
            email_content += "\n\n"

        msg = MIMEMultipart()
        msg['From'] = email_sender
        msg['To'] = recipient
        msg['Subject'] = email_subject

        logger.debug('Email sender: {}'.format(email_sender))
        logger.debug('Email recipient: {}'.format(recipient))

        msg.attach(MIMEText(email_content, 'plain'))

        filename = "jobs.xlsx"
        path_file = os.path.join(os.path.dirname(__file__), 'sheets', filename)
        attachment = open(path_file, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.ehlo()
            server.starttls()
            server.login(gmail_user, gmail_pwd)
            message = msg.as_string()
            server.sendmail(email_sender, email_recipient, message)
            server.close()
            logger.info('Email sent to: {}'.format(email_recipient))
        except:
            logger.error('Email failed to send to: {}'.format(email_recipient), exc_info=True)

