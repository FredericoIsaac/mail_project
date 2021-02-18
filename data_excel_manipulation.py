import pandas


def extract_companies(excel_path, sheet):
    """
    :param excel_path:
    :param sheet:
    :return: Return a Dictionary with info of the company's to send mails
    {
        "contribuinte":
        "identificacao":
        "nome":
        "responsavel":
        "metodo_envio":
        "mail_saft": {"to": "", "cc": ""}
        "saft_submetido":
        "observacao":
    }
    """
    saft_excel = pandas.read_excel(open(excel_path, "rb"), sheet_name=sheet)

    companies_to_send_mail = saft_excel[saft_excel["Enviar Mail"]]

    excel_contribuinte = companies_to_send_mail["NIF's"]
    excel_identificacao = companies_to_send_mail["Nº Emp."]
    excel_nome = companies_to_send_mail["EMPRESAS"]
    excel_responsavel = companies_to_send_mail["Responsável"]
    excel_metodo_envio = companies_to_send_mail["Ficheiro"]
    excel_mail_to = companies_to_send_mail["Mail - To"]
    excel_mail_cc = companies_to_send_mail["Mail - CC"]
    excel_saft_submetido = companies_to_send_mail["Submetido"]
    excel_observacao = companies_to_send_mail["Observações"]

    companies_info = dict()

    for company in range(len(excel_identificacao.values)):
        companies_info[excel_identificacao.values[company]] = {
            "contribuinte": int(excel_contribuinte.values[company]),
            "identificacao": excel_identificacao.values[company],
            "nome": excel_nome.values[company],
            "responsavel": excel_responsavel.values[company],
            "metodo_envio": excel_metodo_envio.values[company],
            "mail_saft": {"to": excel_mail_to.values[company], "cc": excel_mail_cc.values[company]},
            "saft_submetido": excel_saft_submetido.values[company],
            "observacao": excel_observacao.values[company],
        }

    return companies_info
