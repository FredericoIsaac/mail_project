from classes_module import *
from convert_word import *
from data_excel_manipulation import *
from pathlib import Path

WORD_TEMPLATE = "saft_mail_template.docx"
EXCEL_PATH = "Controle de Saft 2021 - Experiencia.xlsx"
SHEET = "Controlo de SAFT V.1.02.02"
ABSOLUTE_PATH_LOGO = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\logo_confere.png"

# Get the corresponding month and year of the SAFT:
month_year = month_in_reference()

# Get the info of excel:
empresas_info = extract_companys(EXCEL_PATH, SHEET)

# Loop trough dict and get every company info:
for key, value in empresas_info.items():
    # Create an Instance of the company:
    company = Company(value)

    # If the SAFT is not submitted send email:
    if not company.saft_submetido:
        # Populate word document with info form company
        completed_word = populate_word(company.nome, company.contribuinte, month_year, WORD_TEMPLATE)

        message = convert_body_to_html(completed_word)

        subject = f"{company.numero} - Saft {str(month_year[0]).zfill(2)}-{month_year[1]}"

        mail = Mail(subject, message, company.mail_saft, ABSOLUTE_PATH_LOGO)

        mail.send_mail()

