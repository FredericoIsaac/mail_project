import pandas as pd


class ExcelMachine:

    def __init__(self, excel_file, sheet=None):
        self.sheet = sheet if sheet else 0
        self.df = pd.read_excel(excel_file, sheet_name=self.sheet)
        self.clients_send_mail = self.to_send_mail()
        self.client_info = self.get_info_clients()

    def to_send_mail(self, column: str = "Enviar Mail"):
        """
        Get the column the true values that are the customers to send mail
        :param column: The column to filter
        :return: Series that has column True
        """
        return self.df[self.df[column]]

    def get_info_clients(self):
        """
        Return a Dictionary with info of the company's to send mails

        :return:{
                10100 : {
                    "Ativo": boolean,
                    "Nº Emp.": str,
                    "EMPRESAS": str,
                    "NIF's": str,
                    "Responsável": str,
                    "Ficheiro" str:
                    "Mail - To": str,
                    "Mail - CC": str,
                    "Mail de Confirmação": boolean,
                    "Observações": str,
                    "Mail Enviado": str,
                    "Submetido": boolean,
                    "Enviar Mail": boolean,
                }
        }
        """
        mails_clients_to_dict = self.clients_send_mail.to_dict('records')
        clients_info_dict = {}

        for index in range(len(mails_clients_to_dict)):
            principal_key = mails_clients_to_dict[index]["Nº Emp."]  # 10120
            value = mails_clients_to_dict[index]
            if principal_key in clients_info_dict:
                numb_store = max(clients_info_dict[principal_key].keys())
                key = numb_store + 1
                clients_info_dict[principal_key].update({key: value})
            else:
                clients_info_dict[principal_key] = {0: value}

        return clients_info_dict

    def introduce_info(self):
        # TODO 1. Criar metodo que coloca informação no excel, caso tenham erro ou caso o saft tenha sido submetido
        pass


excel = ExcelMachine("excel_conference/Controle de Saft 2021 - Experiencia.xlsx")
