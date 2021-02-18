import win32com.client as win32


ABSOLUTE_PATH_SAVE_MAILS = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas" \
                           r"\mail_project\mails_ready\{subject}.msg"


# class Company:
#     def __init__(self, empresa_info):
#         self.metodo_envio = empresa_info["metodo_envio"]
#         self.saft_submetido = empresa_info["saft_submetido"]
#         self.mail_saft = empresa_info["mail_saft"]
#         self.observacao = empresa_info["observacao"]
#
#         self.contribuinte = empresa_info["contribuinte"]
#         self.numero = empresa_info["identificacao"]
#         self.nome = empresa_info["nome"]
#         self.contabilista = empresa_info["responsavel"]


class Mail:
    count = 0

    def __init__(self, subject, html_message, mails: dict, attachment):
        self.attachment = attachment
        self.mail_to = mails["to"]
        self.mail_cc = "" if mails["cc"] != mails["cc"] else mails["cc"]
        self.subject = subject
        self.message_html = html_message

    def send_mail(self, save_mode=True):
        outlook = win32.Dispatch("outlook.application")
        send_mail = outlook.CreateItem(0)
        send_mail.To = self.mail_to
        send_mail.CC = self.mail_cc
        send_mail.Subject = self.subject
        send_mail.HTMLBody = self.message_html

        if self.attachment != r"":
            send_mail.Attachments.Add(self.attachment) if self.attachment else None

        if save_mode:
            send_mail.SaveAs(ABSOLUTE_PATH_SAVE_MAILS.format(subject=self.subject[:5]))
            Mail.count += 1
        else:
            send_mail.Send()
        

