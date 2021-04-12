import os
import win32com.client
import random

path = os.getcwd()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # 6 Is the default inbox folder, it can be changed, e.g. 5 = Sent mailbox
messages = inbox.Items

attachment_list = [] # Used to compare file_name values to previously downloaded files (only current que)


def download_attachments():
    for message in messages:
        if message.Unread:
            print(f"Scanning unread message '{message.Subject}' for excel attachments...")
            attachments = message.Attachments
            for attachment in message.Attachments:
                file_name = attachment.FileName
                renamed_file_name = file_rename(file_name)
                save_file(attachment, renamed_file_name)
                attachment_list.append(renamed_file_name)

                
def file_rename(original_file_name):
    """Renames files that were already downloaded before.
    Takes original file name as input."""
    original_file_name = original_file_name.lower()
    copy_count = 0
    # if ".xls" in original_file_name:  # got rid of this in case someone sends PDF or someone
    if original_file_name in attachment_list:
        while original_file_name in attachment_list:
            copy_count += 1
            original_file_name = f"#{copy_count}#{original_file_name}"
        return original_file_name
    else:
        return original_file_name


def save_file(attachment, new_file_name):
    "Saves file with original or new_name to prevent overrides"
    attachment.SaveAsFile(os.path.join(path, new_file_name))
    print(f"Saving file as {new_file_name}.")
    attachment_list.append(new_file_name)
    print(f"Updated attachment_list is: {attachment_list}.")


input("Please make sure all the emails you want to download are UNREAD.")
input("Please make sure no previous files are in the folder.")

download_attachments()

input(f"{len(attachment_list)} files were downloaded.")
