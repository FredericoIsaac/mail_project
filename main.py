import mail_machine
import excel_machine
import word_machine
import corresponding_date


# Identification of variables in the Program:

# Word:
WORD_TEMPLATE = "./word_template/saft_mail_template.docx"
POPULATED_WORD = "./word_template/Populated_template_V2.docx"

# Excel:
EXCEL_PATH = "excel_conference/Controle de Saft 2021 - Experiencia.xlsx"
SHEET = "Experiencia"


# Get the corresponding month and year of the SAFT:

month_year = corresponding_date.month_in_reference()
month = month_year[0]
year = month_year[1]

# Get the info of excel:

excel_data = excel_machine.ExcelMachine(EXCEL_PATH, SHEET)
companies_data = excel_data.client_info


# Loop trough dict and get every company info:

for n_emp, emp_info in companies_data.items():
    # n_emp = 10101
    # emp_info = { 0: {"Ativo": True,....},{...}...}

    for store, store_info in emp_info.items():
        # Populate Word with info of the company:
        word_transformation = word_machine.WordMachine(
            WORD_TEMPLATE,
            POPULATED_WORD,
            empresa=store_info["EMPRESA"],
            nif=store_info["NIF"],
            id_empresa=n_emp,
        )

        # Get info to send mail
        mail_subject = word_transformation.subject_mail()
        word_transformation.populate_word()
        mail_body = word_transformation.message_to_mail()

        # Send mail
        mail = mail_machine.Mail(mail_subject, mail_body, store_info["Mail - To"], store_info["Mail - CC"])
        mail.send_mail()


print(f"You've sent {mail_machine.Mail.count} e-mails")
