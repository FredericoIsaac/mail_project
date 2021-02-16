import pandas

def extract_companys(excel_path, sheet):
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

    companys_to_send_mail = saft_excel[saft_excel["Mail de Confirmação"] == True]

    excel_contribuinte = companys_to_send_mail["NIF's"]
    excel_identificacao = companys_to_send_mail["Nº Emp."]
    excel_nome = companys_to_send_mail["EMPRESAS"]
    excel_responsavel = companys_to_send_mail["Responsável"]
    excel_metodo_envio = companys_to_send_mail["Ficheiro"]
    excel_mail_to = companys_to_send_mail["Mail - To"]
    excel_mail_cc = companys_to_send_mail["Mail - CC"]
    excel_saft_submetido = companys_to_send_mail["Submetido"]
    excel_obervacao = companys_to_send_mail["Observações"]

    empresas_info = dict()

    for company in range(len(excel_identificacao.values)):
        empresas_info[excel_identificacao.values[company]] = {
                "contribuinte": int(excel_contribuinte.values[company]),
                "identificacao": excel_identificacao.values[company],
                "nome": excel_nome.values[company],
                "responsavel": excel_responsavel.values[company],
                "metodo_envio": excel_metodo_envio.values[company],
                "mail_saft": {"to": excel_mail_to.values[company], "cc": excel_mail_cc.values[company]},
                "saft_submetido": excel_saft_submetido.values[company],
                "observacao": excel_obervacao.values[company],
            }

    return empresas_info