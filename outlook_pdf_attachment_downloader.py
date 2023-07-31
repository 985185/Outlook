import os
import win32com.client

def download_pdf_attachments(folder_path, account_number):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = outlook.Folders.Item(account_number)  # Use Item method to access the desired account
    inbox = account.Folders("Inbox")  # Replace "Inbox" with the folder name if different

    for message in inbox.Items:
        if message.Attachments.Count > 0:
            print(f"Processing email from account {account.Name} with subject: {message.Subject}")
            for attachment in message.Attachments:
                if attachment.FileName.lower().endswith(".pdf"):
                    file_path = os.path.join(folder_path, attachment.FileName)
                    attachment.SaveAsFile(file_path)
                    print(f"Downloaded: {file_path}")

if __name__ == "__main__":
    # Set the folder path where you want to save the PDFs
    download_folder = r"C:\Users\Divija\Documents\Attachments"

    # Create the folder if it doesn't exist
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # Set the correct account index for the desired email account
    account_numbers = [1]  # Update this with the correct account index

    for account_number in account_numbers:
        download_pdf_attachments(download_folder, account_number)
