# Outlook
Certainly! Below is a sample README.md file that you can use to provide information about the Python script that downloads PDF attachments from Outlook emails:

---

# Outlook PDF Attachment Downloader

This Python script allows you to automatically download PDF attachments from Outlook emails. It works with multiple email accounts and saves the downloaded PDF files to a specified folder on your local hard drive.

## Prerequisites

- Python 3.x installed on your system.
- Required Python libraries: `win32com`, `os`

You can install the required libraries using pip:

```bash
pip install pywin32
```

## Setup

1. Clone this repository or download the `outlook_pdf_attachment_downloader.py` script to your local machine.

2. Open the `outlook_pdf_attachment_downloader.py` file in a text editor.

3. Modify the `download_folder` variable to specify the folder where you want to save the downloaded PDF attachments. For example:

   ```python
   download_folder = r"C:\Users\YourUsername\Documents\Attachments"
   ```

   Replace `YourUsername` with your actual username.

4. If you have multiple email accounts in Outlook and want to download PDF attachments from all of them, update the `account_numbers` list to include the correct account indices. To find the correct account indices, run the script with `account_numbers = [0]`, and observe the output. Then update the list with the correct account indices:

   ```python
   account_numbers = [0, 1, 2]  # Update with the correct account indices
   ```

## Usage

1. Make sure Outlook is open and logged in to your email accounts.

2. Open a terminal or command prompt.

3. Navigate to the directory containing the `outlook_pdf_attachment_downloader.py` file.

4. Run the script using the following command:

   ```bash
   python outlook_pdf_attachment_downloader.py
   ```

   The script will start processing emails from the specified accounts' Inboxes, and any PDF attachments found will be downloaded to the specified `download_folder`.

## Note

- This script works with the Windows version of Outlook.

- Use this script responsibly and only on your own Outlook accounts or with proper authorization from the account owners.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

You can create a new file named `README.md` in the same directory as the Python script and paste the above content into it. This README file provides instructions on setup, usage, and any important notes or disclaimers related to the script.

Feel free to customize the README to include any additional information you think might be useful for others who may use or contribute to the project.

If you have any further questions or need more assistance, feel free to ask!
