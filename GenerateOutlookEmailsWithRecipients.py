import win32com.client
import pandas as pd

EXCEL_FILE = pd.read_excel(r"C:\Users\HR-training\Desktop\address_list.xlsx")  # File must be unlocked!

MAIL_SUBJECT = 'Employee Satisfaction Survey 2021 - Results'

MAIL_BODY = """Dear colleagues,

"""


#TODO: Change to OOP?

def send_outlook_mail(recipients, subject='No Subject', body='Blank', send_or_display='Display', copies=None):
    """
    Send an Outlook Text email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: body of the email
    :param send_or_display: Send - send email automatically | Display - email gets created user have to click Send
    :param copies: list of CCs' email addresses
    :return: None
    """
    if len(recipients) > 0 and isinstance(recipient_list, list):  # Checks whether variable "recipient_list" contains more than "0" emails and is the instance of list
        outlook = win32com.client.Dispatch("Outlook.Application")
        ol_msg = outlook.CreateItem(0)

        str_to = ""
        for recipient in recipients:
            # Increases the "str_to" variable by adding the email address and ";" for each recipient
            str_to += recipient + ";"

        ol_msg.To = str_to  # Sets ol_msg attribute .to to generated "str_to" string

        if copies is not None:
            str_cc = ""
            for cc in copies:
                str_cc += cc + ";"  # Same logic as recipients above, but for copies

            ol_msg.CC = str_cc  # Sets ol_msg attribute .cc to generated i

        ol_msg.Subject = subject
        ol_msg.Body = body

        # Allows switching from displaying the outlook message to sending. Recommend to keep as display to confirm everything before sending.
        if send_or_display.upper() == 'SEND':
            ol_msg.Send()
        else:
            # Only open the message (unless we specify in code we want to already send it).
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


if __name__ == '__main__':

    for row in EXCEL_FILE.iterrows():

        print(row
        recipient_list = row

        copies_list = ['simon.soukup@nexentire.com']

        send_outlook_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=MAIL_BODY, send_or_display='Display', copies=copies_list)
