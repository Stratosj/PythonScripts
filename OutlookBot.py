import win32com.client
import pandas as pd
import PySimpleGUI as sg

#TODO: Make it possible to filter addresses by function, team, ...?
#TODO: Make it possible to choose between "Display" or "Send"

### CHANGE SETTINGS HERE (FILE ADDRESS, MAIL SUBJECT, MAIL BODY) ###

# File must be unlocked!
# EXCEL_FILE = pd.read_excel(file_address)
SEND_OR_DISPLAY = "Display"
MAIL_SUBJECT = "Employee Satisfaction Survey 2021 - Results"
MAIL_BODY = """Vážení kolegové,

děkujeme vám a vašim týmům za účast na Průzkumu spokojenosti zaměstnanců. Dalším krokem ke zvýšení spokojenosti zaměstnanců je vytvoření akčního plánu spolu s vašimi členy týmu. Za každý tým navrhněte 1-3 akční kroky, které povedou ke zlepšení spokojenosti. Tyto akční body se musí týkat přímo vašeho týmu, aby je byli schopni všichni členové ovlivnit. Termín každého akčního kroku si stanovíte sami, nejzazší je 31. října 2021. 
Tyto akční plány dejte do přiloženého formuláře a zašlete přes business cooperation report managerovi a team chiefovi.

Součástí tohoto akčního plánu je i možný návrh na zlepšení pro ostatní týmy. Tento návrh na zlepšení musí být dobře popsaný, musí poukázat, co konkrétně zlepší. 

Předem děkujeme za spolupráci,
HR Team
"""

### DO NOT TOUCH ANYTHING BELOW THIS LINE ###


class Messenger():

    def fetch_addresses(self, excel_file_address):
        self.excel_file_address = excel_file_address
        self.recipients_list = []
        self.copies_list = []

        for i, row in self.excel_file_address.iterrows():
            self.recipients_list.append(row["TO"])
            self.copies_list.append(row["CC"])

    def send_outlook_mail(self, subject, body, send_option):
        """
        Send an Outlook Text email to recipients and copy list
        :param recipients: list of recipients' email addresses (list object)
        :param subject: subject of the email
        :param body: body of the email
        :param send_option: Send - send email automatically | Display - email gets created user have to click Send
        :param copies: list of CCs' email addresses
        :return: None
        """
        self.subject = subject
        self.body = body
        self.send_option = send_option

        # Checks whether variable "recipient_list" contains more than "0" emails and is the instance of list
        if len(self.recipients_list) > 0 and isinstance(self.recipients_list, list):
            outlook = win32com.client.Dispatch("Outlook.Application")
            ol_msg = outlook.CreateItem(0)

            str_to = ""
            for recipient in self.recipients_list:
                # Increases the "str_to" variable by adding the email address and ";" for each recipient
                str_to += recipient + ";"

            ol_msg.To = str_to  # Sets ol_msg attribute .to to generated "str_to" string

            if self.copies_list is not None:
                str_cc = ""
                for cc in self.copies_list:
                    str_cc += cc + ";"  # Same logic as recipients above, but for copies

                ol_msg.cc = str_cc  # Sets ol_msg attribute .cc to generated i

            ol_msg.Subject = self.subject
            ol_msg.Body = self.body

            # Allows switching from displaying the outlook message to sending. Recommend to keep as display to confirm everything before sending.
            if self.send_option.upper() == "SEND":
                ol_msg.Send()
            else:
                # Only open the message (unless we specify in code we want to already send it).
                ol_msg.Display()
        else:
            print("Recipient email address - NOT FOUND")


if __name__ == "__main__":

    layout = [
        [sg.Text("Select file with e-mail addresses.")],
        [sg.Input(), sg.FileBrowse('FileBrowse', file_types=(("Excel files only", "*.xlsx"),))],
        [sg.Submit(), sg.Cancel()],
    ]

    window = sg.Window('Outlook Message Bot', layout)

    while True:
        event, values = window.read()
        if event is None or event == 'Cancel':
            break

        if event == 'Submit':
            file_address = values['FileBrowse']
            file_address = pd.read_excel(file_address)
            MessageBot = Messenger()
            MessageBot.fetch_addresses(excel_file_address=file_address)
            MessageBot.send_outlook_mail(subject=MAIL_SUBJECT, body=MAIL_BODY, send_option=SEND_OR_DISPLAY)
            window.close()
