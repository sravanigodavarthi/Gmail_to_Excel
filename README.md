# Gmail_to_Excel
This Python script allows you to extract specific email messages from your Gmail inbox using IMAPClient, retrieve their subject and content, and save the data into an Excel file. 
Additionally, it can create a label/tag in your Gmail account and move the selected emails to that label.
## The reason behind this script
I receive helpful Python tricks emails from Real Python, a great website for improving Python skills. However, it's not always convenient to read and use the tricks immediately after receiving the email. So, I decided to automate the process. I created a script that extracts the content of the tricks from the gmail and stores them in an Excel file. Now, whenever I want to refer to those tricks, I can simply open the Excel file and easily find the information I need. I can even add comments or notes alongside the tricks in the Excel file to help me remember important details. This way, I have all the tricks organized and easily accessible whenever I want to review or use them again.
## Prerequisites
 * Python 3.x.
 * pandas library.
 * IMAPClient library:
      The IMAPClient library is a widely used Python library for Gmail. It allows you to connect to a Gmail account over IMAP (Internet Message Access Protocol) and perform various operations such as retrieving emails, searching for specific messages, fetching attachments,         and perform other operations. For more details, please refer to the official documentation [IMAPClient Documentation](https://imapclient.readthedocs.io/en/master/#imapclient.IMAPClient.search).
## Installation
1. Make sure you have Python 3.x and the package manager pip installed.
2. To install the package in your desired Python environment or interpreter, use the following command:

   pip install git+https://github.com/sravanigodavarthi/Gmail-to-Excel.git
## Usage
You can use the package in your projects or scripts as needed, as shown in the code snippet below:

```bash
# Import the package
import Gmail_to_excel_Package

# Configuration data for retrieving Gmail messages and writing to Excel
config = {
    'username': 'YOUR_GMAIL_USERNAME',
    'password': 'YOUR_APP_PASSWORD',
    'search_query': 'YOUR_SEARCH_QUERY',
    'output_file_path': 'YOUR_OUTPUT_EXCEL_PATH',
    'gmail_label_name': 'YOUR_GMAIL_LABEL_NAME'
}

# Call the function to retrieve Gmail messages and write to Excel
Gmail_to_excel_Package.write_gmail_messages_to_excel(config)
```
### Configuration Data

Under the `config` dictionary, replace the placeholder values with your actual information:

- `'YOUR_GMAIL_USERNAME'`: Your Gmail username.
  
- `'YOUR_APP_PASSWORD'`: Your application-specific password (generated from your Google account settings). To ensure the security of your Gmail account, it is recommended to generate an application-specific password instead of using your regular Gmail password. To generate an application-specific password, you can visit the following link: [Generate Application-Specific Password](https://support.google.com/accounts/answer/185833).
  
- `'YOUR_SEARCH_QUERY'`: Your desired search query to filter emails. In my case, the search query `'SUBJECT "[*PyTricks]: "'` is set to filter emails based on the subject containing `[*PyTricks]:`.
  
- `'YOUR_OUTPUT_EXCEL_PATH'`: The desired path for the output Excel file.
  
- `'YOUR_GMAIL_LABEL_NAME'`: The desired Gmail label name.

Make sure to replace these values with your specific information before running the code.

## Output
1. The script will fetch emails matching the search query from your Gmail account.

2. The subject and content of the emails will be extracted and stored in an Excel file.

3. If the Excel file already exists, the script will append the new data to it.

4. If the specified label/tag already exists in your Gmail account, the script will move the selected emails to that label.

5. If the label/tag doesn't exist, the script will create it and move the emails to that label.

## Notes
* It's recommended to run the script periodically to keep the Excel file up to date with new emails.
