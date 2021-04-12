import os
import win32com.client
import random

path = os.getcwd()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(5) # 6 Is the default inbox folder, it can be changed, e.g. 5 = Sent mailbox
messages = inbox.Items
attachment_list = []

# TODO: Change to functions and clean up this mess.
# TODO: Giver user ability to select folder?

def saveattachments():
    for message in messages:
        if message.Unread:
            print(f"Scanning unread message '{message.Subject}' for excel attachments...")
            attachments = message.Attachments
            for attachment in message.Attachments:
                file_name = str(attachment.FileName)
                file_name = file_name.lower()
                if ".xlsx" in file_name:
                    # print(f"Found {attachment}.")
                    # print(f"Scanning attachment_list to see if {attachment} was already downloaded.")
                    if file_name in attachment_list:
                        new_file_name = file_name
                        copy_count = 0
                        while new_file_name in attachment_list:
                            copy_count += 1
                            # print(f"The file {new_file_name} was already downloaded before.")
                            new_file_name = f"#{copy_count}#{new_file_name}"
                            # print(f"Adding {copy_count} in front.")
                        attachment.SaveAsFile(os.path.join(path, str(new_file_name)))
                        print(f"Saving file as {new_file_name}.")
                        attachment_list.append(new_file_name)
                        print(f"Updated attachment_list is: {attachment_list}.")
                    elif file_name not in attachment_list:
                        # print(f"The file {attachment} was not downloaded before.")
                        attachment.SaveAsFile(os.path.join(path, str(file_name)))
                        print(f"Saving file as {file_name}.")
                        attachment_list.append(file_name)
                        print(f"Updated attachment_list is: {attachment_list}.")


input("Please make sure all the emails you want to download are UNREAD.")
input("Please make sure no previous files are in the folder.")

saveattachments()

input(f"{len(attachment_list)} files were downloaded.")
