import corresponding_date
from mailmerge import MailMerge
import mammoth
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import docx2txt


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


class WordMachine:

    def __init__(self, template_path: str, output_path: str, **kwargs):

        # Date
        month_year = corresponding_date.month_in_reference()
        self.month = month_year[0]
        self.year = month_year[1]
        self.extended_month = MONTH_DICT[self.month]

        # Paths
        self.template_path = template_path
        self.output_path = output_path

        # Data to fill fields
        self.company = kwargs.get("empresa", "")
        self.nif = str(kwargs.get("nif", ""))
        self.id_company = kwargs.get("id_empresa", "")

    def populate_word(self):
        """
        Populate the word Document and save to a new file
        """
        document = MailMerge(self.template_path)

        # Get the name of the fields in Word
        # fields = document.get_merge_fields()
        # for field in document.get_merge_fields():
        #     print(field)

        document.merge(
            empresa=self.company,
            ano_referente=str(self.year),
            mes_referente=str(self.extended_month),
            nif=str(self.nif),
        )

        document.write(self.output_path)

    def word_to_html(self):
        """
        :return: An HTML string with image of the signature Frederico Gago
        """
        with open(self.output_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value  # The generated HTML
            # messages = result.messages  # Any messages, such as warnings during conversion

        # Add Signature:
        html += f"<br><br><br><br><img src='cid:logo'>"

        return html

    def message_to_mail(self):
        # Encapsulate the plain and HTML versions of the message body in an
        # 'alternative' part, so message agents can decide which they want to display.
        message = MIMEMultipart("alternative")

        # Get data from Word
        text = docx2txt.process(self.output_path)

        # Transform Word to HTML
        html = self.word_to_html()

        plain_text = MIMEText(text, "plain")
        message.attach(plain_text)

        html_text = MIMEText(html, "html")
        message.attach(html_text)

        # We reference the image in the IMG SRC attribute by the ID we give it <img src='cid:logo'>
        with open("images/signature.png", "rb") as img:
            signature = MIMEImage(img.read())
            signature.add_header("Content-ID", "<logo>")
            message.attach(signature)

        return message

    def subject_mail(self):
        return f"{self.id_company} - Saft {str(self.month).zfill(2)}-{self.year}"
