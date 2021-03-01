# MAIL MACHINE


# def mail(subject, body, to, cc=None):
#
#     to = to
#     subject = subject
#     body = body
#     cc = "" if cc is None else cc
#
#     msgRoot = MIMEMultipart("related")
#     msgRoot["Subject"] = subject
#     msgRoot["From"] = MAIL_FROM
#     msgRoot["To"] = to
#     msgRoot.preamble = "Envio do SAFT"
#
#     message = word_machine.word.message_to_mail()
#     msgRoot.attach(message)
#
#     try:
#         smtp = smtplib.SMTP()
#         smtp.connect(SERVER)
#         smtp.login(MAIL_FROM, PASSWORD_MAIL)
#         smtp.sendmail(MAIL_FROM, to, msgRoot.as_string())
#         smtp.quit()
#     except smtplib.SMTPException:
#         print(f"Error: unable to send email to {subject[:5]}:{to}")


# class Mail:
#     count = 0
#
#     def __init__(self, subject, html_message, mails: dict, attachment):
#         self.attachment = attachment
#         self.mail_to = mails["to"]
#         self.mail_cc = "" if mails["cc"] != mails["cc"] else mails["cc"]
#         self.subject = subject
#         self.message_html = html_message
#
#     def send_mail(self, save_mode=True):
#         outlook = win32.Dispatch("outlook.application")
#         send_mail = outlook.CreateItem(0)
#         send_mail.To = self.mail_to
#         send_mail.CC = self.mail_cc
#         send_mail.Subject = self.subject
#         send_mail.HTMLBody = self.message_html
#
#         if self.attachment != r"":
#             send_mail.Attachments.Add(self.attachment) if self.attachment else None
#
#         if save_mode:
#             send_mail.SaveAs(ABSOLUTE_PATH_SAVE_MAILS.format(subject=self.subject[:5]))
#             Mail.count += 1
#         else:
#             send_mail.Send()


# CONVERT WORD

# from __future__ import print_function
# from mailmerge import MailMerge
# import mammoth
#
# POPULATED_WORD = "./word_template/Populated_template_V2.docx"
# ABSOLUTE_PATH_LOGO = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\signature.png"
#
# MONTH_DICT = {
#     1: 'Janeiro',
#     2: 'Fevereiro',
#     3: 'Março',
#     4: 'Abril',
#     5: 'Maio',
#     6: 'Junho',
#     7: 'Julho',
#     8: 'Agosto',
#     9: 'Setembro',
#     10: 'Outubro',
#     11: 'Novembro',
#     12: 'Dezembro',
# }
#
#
#
#
# def populate_word(company_name, contribuinte, corresponding_date: tuple, path_template):
#
#     document = MailMerge(path_template)
#     # print(document.get_merge_fields())
#     document.merge(
#         empresa=str(company_name),
#         ano_referente=str(corresponding_date[1]),
#         mes_referente=str(MONTH_DICT[corresponding_date[0]]),
#         nif=str(contribuinte)
#     )
#
#     document.write(POPULATED_WORD)
#     return POPULATED_WORD
#
#
# def convert_body_to_html(path_template) -> str:
#     """
#     :return: An HTML string with image of the assignature Frederico Gago
#     """
#
#     with open(path_template, "rb") as docx_file:
#         result = mammoth.convert_to_html(docx_file)
#         html = result.value  # The generated HTML
#         messages = result.messages  # Any messages, such as warnings during conversion
#
#     html = html + f"""<br><br><br><br><img src="{ABSOLUTE_PATH_LOGO}"
#     alt="Com os melhores Cumprimentos,\n Frederico Gago\n Confere - Silva & Sabino">"""
#
#     return html


# DATA EXCEL MANIPULATION

# import pandas
#
#
# def extract_companies(excel_path, sheet):
#     """
#     :param excel_path:
#     :param sheet:
#     :return: Return a Dictionary with info of the company's to send mails
#     {
#         "contribuinte":
#         "identificacao":
#         "nome":
#         "responsavel":
#         "metodo_envio":
#         "mail_saft": {"to": "", "cc": ""}
#         "saft_submetido":
#         "observacao":
#     }
#     """
#     saft_excel = pandas.read_excel(open(excel_path, "rb"), sheet_name=sheet)
#
#     companies_to_send_mail = saft_excel[saft_excel["Enviar Mail"]]
#
#     excel_contribuinte = companies_to_send_mail["NIF's"]
#     excel_identificacao = companies_to_send_mail["Nº Emp."]
#     excel_nome = companies_to_send_mail["EMPRESAS"]
#     excel_responsavel = companies_to_send_mail["Responsável"]
#     excel_metodo_envio = companies_to_send_mail["Ficheiro"]
#     excel_mail_to = companies_to_send_mail["Mail - To"]
#     excel_mail_cc = companies_to_send_mail["Mail - CC"]
#     excel_saft_submetido = companies_to_send_mail["Submetido"]
#     excel_observacao = companies_to_send_mail["Observações"]
#
#     companies_info = dict()
#
#     for company in range(len(excel_identificacao.values)):
#         companies_info[excel_identificacao.values[company]] = {
#             "contribuinte": int(excel_contribuinte.values[company]),
#             "identificacao": excel_identificacao.values[company],
#             "nome": excel_nome.values[company],
#             "responsavel": excel_responsavel.values[company],
#             "metodo_envio": excel_metodo_envio.values[company],
#             "mail_saft": {"to": excel_mail_to.values[company], "cc": excel_mail_cc.values[company]},
#             "saft_submetido": excel_saft_submetido.values[company],
#             "observacao": excel_observacao.values[company],
#         }
#
#     return companies_info


# main


# Get the info of excel:
# from data_excel_manipulation import *
# companies_info = extract_companies(EXCEL_PATH, SHEET)


# completed_word = populate_word(value["nome"], value["contribuinte"], month_year, WORD_TEMPLATE)
# message = convert_body_to_html(completed_word)


# print(key, value)

# company = Company(value)
# Populate word document with info form company
# completed_word = populate_word(company.nome, company.contribuinte, month_year, WORD_TEMPLATE)
# message = convert_body_to_html(completed_word)
#
# subject = f"{company.numero} - Saft {str(month_year[0]).zfill(2)}-{month_year[1]}"
#
# mail = Mail(subject, message, company.mail_saft, ABSOLUTE_PATH_LOGO)
# mail.send_mail()