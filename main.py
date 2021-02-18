from mail_sender import *
from convert_word import *
from data_excel_manipulation import *


WORD_TEMPLATE = "./word_template/saft_mail_template.docx"
EXCEL_PATH = "excel_conference/Controle de Saft 2021 - Experiencia.xlsx"
SHEET = "Controlo de SAFT V.1.02.02"
ABSOLUTE_PATH_ATTACHMENT = r""

# Get the corresponding month and year of the SAFT:
month_year = month_in_reference()

# Get the info of excel:
companies_info = extract_companies(EXCEL_PATH, SHEET)

# Loop trough dict and get every company info:
for value in companies_info.values():
    # Populate Word with info of the company:
    completed_word = populate_word(value["nome"], value["contribuinte"], month_year, WORD_TEMPLATE)
    message = convert_body_to_html(completed_word)

    # Send mail
    subject = f"{value['identificacao']} - Saft {str(month_year[0]).zfill(2)}-{month_year[1]}"
    mail = Mail(subject, message, value["mail_saft"], ABSOLUTE_PATH_ATTACHMENT)
    mail.send_mail()

    # company = Company(value)
    # Populate word document with info form company
    # completed_word = populate_word(company.nome, company.contribuinte, month_year, WORD_TEMPLATE)
    # message = convert_body_to_html(completed_word)
    #
    # subject = f"{company.numero} - Saft {str(month_year[0]).zfill(2)}-{month_year[1]}"
    #
    # mail = Mail(subject, message, company.mail_saft, ABSOLUTE_PATH_LOGO)
    # mail.send_mail()


print(f"You've sent {Mail.count} e-mails")
