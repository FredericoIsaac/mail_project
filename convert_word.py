from __future__ import print_function
from mailmerge import MailMerge
import mammoth
import datetime

POPULATED_WORD = "./word_template/Populated_template_V2.docx"
ABSOLUTE_PATH_LOGO = r"C:\Users\Frederico\Desktop\Frederico Gago\Confere\Programas\mail_project\logo_confere.png"
MONTH_DICT = {
    1: 'Janeiro',
    2: 'Fevereiro',
    3: 'Março',
    4: 'Abril',
    5: 'Maio',
    6: 'Junho',
    7: 'Julho',
    8: 'Agosto',
    9: 'Setembro',
    10: 'Outubro',
    11: 'Novembro',
    12: 'Dezembro',
}


def month_in_reference():
    """
    :return a Tuple of the corresponding Month and Year of SAFT
    Example: current month 1 (January) of 2021 returns 12 (December) of 2020
    """
    months = [n for n in range(1, 13)]
    current_date = datetime.date.today()
    current_month = current_date.timetuple()[1]
    last_month = months[current_month - 2]
    current_year = current_date.timetuple()[0]

    if last_month == 12:
        year = current_year - 1
    else:
        year = current_year

    return last_month, year


def populate_word(company_name, contribuinte, corresponding_date: tuple, path_template):

    document = MailMerge(path_template)
    # print(document.get_merge_fields())
    document.merge(
        empresa=str(company_name),
        ano_referente=str(corresponding_date[1]),
        mes_referente=str(MONTH_DICT[corresponding_date[0]]),
        nif=str(contribuinte)
    )

    document.write(POPULATED_WORD)
    return POPULATED_WORD


def convert_body_to_html(path_template) -> str:
    """
    :return: An HTML string with image of the assignature Frederico Gago
    """

    with open(path_template, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # The generated HTML
        messages = result.messages  # Any messages, such as warnings during conversion

    html = html + f"""<br><br><br><br><img src="{ABSOLUTE_PATH_LOGO}"
    alt="Com os melhores Cumprimentos,\n Frederico Gago\n Confere - Silva & Sabino">"""

    return html

