import email
import sys
import os

import pandas as pd
from imapclient import IMAPClient
from imapclient.exceptions import LoginError


def write_gmail_messages_to_excel(config_dict):

    # IMAPClient connection setup
    username = config_dict.get('username')
    password = config_dict.get('password')
    with IMAPClient('imap.gmail.com') as client:
        try:
            client.login(username, password)
        except LoginError as e:
            print("Login failed:", str(e))
            sys.exit()

        # Select the desired folder
        client.select_folder('INBOX')

        # Define the search query for specific messages
        search_query = config_dict.get('search_query')

        # Search for messages matching the search query
        messages = client.search(search_query)

        if messages:
            # Initialize the dictionary to store message data
            message_data_dict = {}

            # Fetch and process the message data
            message_data_dict = fetch_message_data(messages, client)

            # Write the fetched message data to an Excel file
            output_file_path = config_dict.get('output_file_path')
            write_to_excel(message_data_dict, output_file_path)

            # Add a Gmail label to the specified messages
            gmail_label_name = config_dict.get('gmail_label_name')
            add_gmail_label(messages, gmail_label_name, client)
        else:
            print(f"No messages matching the search query {search_query} were found in the INBOX folder.")
            sys.exit()


def fetch_message_data(message_data, client):
    # Fetches the message data
    try:
        data_dict = {}

        # Fetch the message content from the server
        message_content = client.fetch(message_data, ['BODY[]', 'BODY[HEADER.FIELDS (SUBJECT)]'])

        for gmail_id, data in message_content.items():
            # Extract subject title
            subject = data[b'BODY[HEADER.FIELDS (SUBJECT)]'].decode()
            subject_title = subject.split("?=")[1]

            # Extract message body
            message_body = email.message_from_bytes(data[b'BODY[]'])

            # Extract content
            if message_body.is_multipart():
                for part in message_body.get_payload():
                    if part.get_content_type() == 'text/plain':
                        content = part.get_payload(decode=True)
                        break
                else:
                    content = message_body.get_payload(decode=True)

            # Decode content and extract main content
            content = content.decode('utf-8')
            main_content = content.split("-----")[0]

            # Store subject title and main content in data dictionary
            data_dict[subject_title] = main_content
            return data_dict
    except ValueError as e:
        print(str(e))
        sys.exit()


def add_gmail_label(message_list, label_name, client):
    # Adds a Gmail label to the specified messages

    if client.folder_exists(label_name):
        # Move messages to an existing folder
        client.move(message_list, label_name)
        print(f"Gmail's moved to tag {label_name}")
    else:
        # Create a new folder and move messages
        client.create_folder('Real_python_tricks')
        client.move(message_list, 'Real_python_tricks')
        print(f"Gmail's moved to tag {label_name}")


def write_to_excel(write_data, file_path):
    # Writes the fetched message data to an Excel file

    # Prepare the data for the DataFrame
    rows = [{'Subject': subject, 'Code': code} for subject, code in write_data.items()]
    df = pd.DataFrame(rows)

    try:
        if os.path.exists(file_path):
            # If the file exists, append the data to it
            df_existing = pd.read_excel(file_path)
            df_updated = pd.concat([df_existing, df], ignore_index=True)
            df_updated.to_excel(file_path, index=False)
            print("Excel write operation successful")
        else:
            # If the file doesn't exist, create a new file and write the data
            df.to_excel(file_path, index=False)

            print("Excel write operation successful")
    except (PermissionError, ValueError) as e:
        print(str(e))
        print(f"Please make sure the file path {file_path} is correct and file have the necessary permissions.")
        sys.exit()


if __name__ == "__main__":
    # Description of the script
    """
    This script retrieves specific email messages from Gmail using IMAPClient,
    extracts relevant data, and writes it to an Excel file.
    """
    config = {
        'username': 'YOUR_GMAIL_USERNAME',
        'password': 'YOUR_APP_PASSWORD',
        'search_query': 'YOUR_SEARCH_QUERY',
        'output_file_path': 'YOUR_OUTPUT_EXCEL_PATH',
        'gmail_label_name': 'YOUR_GMAIL_LABEL_NAME'
        }
    write_gmail_messages_to_excel(config)