# --------------------------- IMPORTS --------------------------- #


import mail_machine
import excel_machine
import word_machine
import corresponding_date
from prettytable import PrettyTable


# --------------------------- CONSTANT VARIABLES --------------------------- #


# Word:
WORD_TEMPLATE = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\word_template\saft_mail_template.docx"
POPULATED_WORD = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\word_template\Populated_template_V2.docx"

# Excel:
EXCEL_PATH = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\excel_conference\Controle de Saft 2021 - Experiencia.xlsx"
SHEET = "Experiencia"

# --------------------------- DATE PAST MONTH --------------------------- #


month_year = corresponding_date.month_in_reference()
month = month_year[0]
year = month_year[1]

# --------------------------- EXCEL EXTRACTOR --------------------------- #


def excel_extractor():
    excel_data = excel_machine.ExcelMachine(EXCEL_PATH, SHEET)
    return excel_data.client_info


# --------------------------- EXCEL Printer --------------------------- #


def print_excel_list(company_info):
    """
    Print a pretty Table with the info of the company's to send email.
    :param company_info: Dict of all the company info
    """

    show_clients_list = PrettyTable(["Store", "Nº", "Company", "NIF", "Mail-to"])
    for company in company_info.values():
        for store, info in company.items():
            show_clients_list.add_row([store, info["Nº Emp."], info["EMPRESA"], info["NIF"], info["Mail - To"]])

    print(show_clients_list)


# --------------------------- Save & Send --------------------------- #

def send_save_mail(data, mode):
    for company in data.values():
        for store, info in company.items():

            company_name = info["EMPRESA"]
            id_company = info["Nº Emp."]
            nif_company = info["NIF"]
            mail_to = info["Mail - To"]
            mail_cc = info["Mail - CC"]
            sent_mails_list = mail_machine.Mail.sent_mails

            # Populate Word with info of the company:
            word_transformation = word_machine.WordMachine(
                WORD_TEMPLATE,
                POPULATED_WORD,
                empresa=company_name,
                nif=nif_company,
                id_empresa=id_company,
            )

            # Get info to send mail
            mail_subject = word_transformation.subject_mail()
            word_transformation.populate_word()

            if mode == "save":
                body_save = word_transformation.word_to_html()
                mail = mail_machine.Mail(mail_subject, body_save, mail_to, mail_cc)
                mail.save_mails()
            elif mode == "send":

                # Checks if already sent mail to the same email with the same company name
                if id_company in [elem for sublist in sent_mails_list for elem in sublist] and \
                        mail_to in [elem for sublist in sent_mails_list for elem in sublist]:
                    continue
                else:
                    sent_mails_list.append([id_company, mail_to])

                body_send = word_transformation.message_to_mail()
                mail = mail_machine.Mail(mail_subject, body_send, mail_to, mail_cc)
                mail.send_mail()

                # take note in excel
                # Create Instance
                input_excel = excel_machine.ExcelMachine(EXCEL_PATH, SHEET)
                input_excel.introduce_info(company_number=int(id_company), store_number=int(store), info_saft="mail_enviado")


# --------------------------- PROGRAM FLOW --------------------------- #


menu_string = """Menu
1 - Imprimir lista de clientes com o SAFT pendente
2 - Enviar Mail's
0 - Exit
"""

companies_data = None

while True:
    menu_choice = input(menu_string)

    if menu_choice == "0":
        exit()
    elif menu_choice == "1":
        companies_data = excel_extractor()
        print_excel_list(companies_data)
    elif menu_choice == "2" and companies_data:
        send_save_mail(companies_data, "save")

        check_mails = input(f"Mail have been saved to {mail_machine.ABSOLUTE_PATH_SAVE_MAILS[:84]},"
                            f" please check if everything ok!\nWrite 'Enviar Mail' to send mails\n"
                            f"To Cancel and return to menu press any other key:\n")

        if check_mails.lower() == "enviar mail":
            send_save_mail(companies_data, "send")
            print(f"You've sent {mail_machine.Mail.count} e-mails")
    else:
        print("First choose option 1 to view data\n")
