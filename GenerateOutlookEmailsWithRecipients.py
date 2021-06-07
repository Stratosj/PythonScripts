import win32com.client
import pandas as pd

# TODO: Make this WINFORM (winform tutorial.py / https://www.youtube.com/watch?v=RFwTk4twaOI&ab_channel=theI.T.Guy)
# TODO: Strip this of any unnecessary things and make sure you understand what's going on in the code.


# # excel_file = pd.read_excel(os.path.abspath(r"C:\Users\HR-training\Documents\Python\AKTION RAW 210101-210131 (All Teams) v1.0.xlsx"))
EXCEL_FILE = pd.read_excel(r"D:\NEXEN\04 Controlling\HR Dashboard\Simon\HR Dashboard\Source\AKTION RAW 210101-210131 (All Teams) v1.0.xlsx") # File must be unlocked!

# first_filter_choice = input("What teams shall I filter?\n") # Choses which column to filter
# first_criteria = input(f"What positions shall I filter {first_filter_choice} for?\n")
# first_criteria_list = first_criteria.split(",") # TODO: Add a while loop.
# secondary_filter_choice = input("What superiors shall I filter?\n") # Choses which column to filter


# # https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html
# position_filtered = EXCEL_FILE.loc[first_criteria_list]
# # FIXME: This doesn't work (should filter by more than one criteria)
# # FIXME: Exception has occurred: ValueError - ('Lengths must match to compare', (959,), (2,))

# # Filters pandas data frame by required filter
# position_filtered = position_filtered.loc[position_filtered[secondary_filter_choice] == second_criteria]
# position_filtered = position_filtered["Osobní číslo"].tolist()  # Transforms DF column results to a list to be used in outlook

# print(position_filtered)
# # position_filtered = position_filtered["Osobní číslo"].tolist()

# # print(position_filtered)


################################################################## OUTLOOK MESSAGE ##################################################################


# Hard coded email subject
MAIL_SUBJECT = 'Předmět zprávy'

# Hard coded email text
MAIL_BODY = 'Vážení kolegové,'


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


# RECIPIENT_DATABASE = [['jiri.stros@nexentire.com', 'simon.soukup@nexentire.com'],
#                       ['stratosj@gmail.com', 'daniel.fokt@nexentire.com']]

if __name__ == '__main__':

    for row in EXCEL_FILE.iterrows():

        print(row.
        # recipient_list = row

        # copies_list = ['simon.soukup@nexentire.com']

        # send_outlook_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=MAIL_BODY, send_or_display='Display',
        #                   copies=copies_list)
