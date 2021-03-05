from os.path import basename
from os.path import abspath
import win32com.client as win32
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
import os

# Mail Variables:
MAIL_FROM = "fredericogago@confere.pt"
SERVER = "mail.confere.pt"
PASSWORD_MAIL = os.environ["PASSWORD_MAIL"]

ABSOLUTE_PATH_SAVE_MAILS = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\mails_ready\{subject}.msg"
ATTACHMENTS = [r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\images\signature.png"]


class Mail:
    count = 0
    sent_mails = []

    def __init__(self, subject, body, to, cc=None, attachment=None):
        self.to = to
        self.subject = subject
        self.body = body
        self.cc = "" if cc != cc else cc
        self.attachment_list = ATTACHMENTS if attachment is None else ATTACHMENTS.append(attachment)

    def send_mail(self):

        receivers = [self.to] + self.cc.split(";")

        msg_root = MIMEMultipart("related")
        msg_root["Subject"] = self.subject
        msg_root["From"] = MAIL_FROM
        msg_root["To"] = self.to
        msg_root["Cc"] = self.cc
        msg_root.preamble = self.subject

        message = self.body
        msg_root.attach(message)

        # Attach all the path in the list
        for path in self.attachment_list:
            with open(path, "rb") as file:
                part = MIMEApplication(file.read(), Name=basename(path))
            part["Content-Disposition"] = f"attachment; filename={basename(path)}"
            msg_root.attach(part)

        try:
            smtp = smtplib.SMTP()
            smtp.connect(SERVER)
            smtp.login(MAIL_FROM, PASSWORD_MAIL)
            smtp.sendmail(MAIL_FROM, receivers, msg_root.as_string())
            smtp.quit()
        except smtplib.SMTPException:
            print(f"Error: unable to send email to {self.subject[:5]}: {self.to}")

    def save_mails(self):
        outlook = win32.Dispatch("outlook.application")
        save_mail = outlook.CreateItem(0)
        save_mail.To = self.to
        save_mail.CC = self.cc
        save_mail.Subject = self.subject
        save_mail.HTMLBody = self.body

        for str_path in self.attachment_list:
            save_mail.Attachments.Add(abspath(str_path))

        save_mail.SaveAs(ABSOLUTE_PATH_SAVE_MAILS.format(subject=self.subject[:5]))


# TODO 2. Fazer com que nao envie dois mails iguais
