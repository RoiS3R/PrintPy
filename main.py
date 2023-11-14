# Import necessary libraries
import time
import imaplib
import email
import os
import cups
import logging
from dotenv import load_dotenv
import email

# Load environment variables
load_dotenv()
PRINT_COMMAND = os.getenv('PRINT_COMMAND')

# Setup basic logging
logging.basicConfig(filename='email_to_print.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Constants for email configuration
SMTP_USER = os.getenv('SMTP_USER')
SMTP_PASS = os.getenv('SMTP_PASS')
SMTP_HOST = os.getenv('SMTP_HOST')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
SMTP_TLS = os.getenv('SMTP_TLS') == 'true'

REFRESH_INTERVAL = int(os.getenv('REFRESH_INTERVAL'))

# Function to print a PDF file


def print_pdf(printer_url, pdf_file_path, color_mode='bw', duplex='false', media='A4', copies=1, page_ranges=None):
    """
    Prints a PDF file to a specified printer using CUPS.

    Args:
        printer_url (str): The URL of the printer to use.
        pdf_file_path (str): The path to the PDF file to print.
        color_mode (str, optional): The color mode to use for printing. Defaults to 'bw'.
        duplex (str, optional): Whether to print double-sided or not. Defaults to 'false'.
        media (str, optional): The media size to use for printing. Defaults to 'A4'.
        copies (int, optional): The number of copies to print. Defaults to 1.
        page_ranges (str, optional): The page ranges to print. Defaults to None.

    Returns:
        None
    """
    # Printer configuration
    printer_uri = "dnssd://Kyocera%20TASKalfa%204054ci._ipps._tcp.local./?uuid=4509a320-00ff-00b4-0059-002507530fcd"
    printer_name = "Kyocera_TASKalfa_4054ci"

    # Ensure that the file exists
    if not os.path.exists(pdf_file_path):
        logging.error(f"File {pdf_file_path} does not exist.")
        return

    # Create a connection to the CUPS server
    conn = cups.Connection()

    # Get a list of printers to verify the printer is available
    printers = conn.getPrinters()

    # Printing options
    options = {
        'media': media,
        'copies': str(copies),
        'fit-to-page': 'True',
        'sides': 'two-sided-long-edge' if duplex == 'true' else 'one-sided',
        'color-mode': 'color' if color_mode == 'color' else 'monochrome',
    }

    if page_ranges:
        options['page-ranges'] = page_ranges

    # Check if the specified printer is in the list of printers
    if printer_name not in printers:
        logging.error(f"Printer {printer_name} not found.")
        return

    # Print the file
    print_job_id = conn.printFile(
        printer_name, pdf_file_path, "Python Print Job", options)
    logging.info(f"Print job sent. Job ID: {print_job_id}")

# Function to save a file to the local file system, handling duplicates


def save_file_to_local(directory, filename, data):
    """
    Saves a file to the specified directory with the given filename. If a file with the same name already exists,
    a counter is appended to the filename to avoid overwriting the existing file.

    Args:
        directory (str): The directory to save the file in.
        filename (str): The name to give the file.
        data (bytes): The data to write to the file.

    Returns:
        str: The full path to the saved file.
    """
    if not os.path.exists(directory):
        os.makedirs(directory)

    file_path = os.path.join(directory, filename)
    counter = 1

    # Handle duplicate filenames
    while os.path.exists(file_path):
        base, extension = os.path.splitext(filename)
        file_path = os.path.join(directory, f"{base}({counter}){extension}")
        counter += 1

    with open(file_path, 'wb') as file:
        file.write(data)

    logging.info(f"File saved: {file_path}")
    return file_path

# Function to delete a file from local storage


def delete_file(directory, filename):
    """
    Deletes a file with the given filename in the specified directory.

    Args:
        directory (str): The directory where the file is located.
        filename (str): The name of the file to be deleted.

    Returns:
        None
    """
    file_path = os.path.join(directory, filename)
    if os.path.exists(file_path):
        os.remove(file_path)
        logging.info(f"Deleted file: {file_path}")
    else:
        logging.error(f"File not found for deletion: {file_path}")

# Function to determine if the print should be black and white


def is_black_and_white(subject):
    """
    Determines if an email subject contains the "-c=color" flag, indicating that the email should be printed in color.

    Args:
        subject (str): The subject line of the email.

    Returns:
        bool: True if the email should be printed in black and white, False if it should be printed in color.
    """
    return "-c=color" not in subject

# Function to determine if the print should be duplex


def is_duplex(subject):
    """
    Check if the subject of an email contains the string "-d=duplex".

    Args:
        subject (str): The subject of the email.

    Returns:
        bool: True if the subject contains "-d=duplex", False otherwise.
    """
    return "-d=duplex" in subject

# Function to set the number of copies


def get_copies(subject):
    """
    Parses the subject line of an email to extract the number of copies to be printed.

    Args:
        subject (str): The subject line of the email.

    Returns:
        int: The number of copies to be printed. Defaults to 1 if no '-copies=' parameter is found in the subject line.
    """
    copies = 1
    if "-copies=" in subject:
        copies = int(subject.split("-copies=")[1].split()[0])
    return copies

# Function to set the page ranges


def get_page_ranges(subject):
    """
    Extracts the page ranges from the subject of an email.

    Args:
        subject (str): The subject of the email.

    Returns:
        str: The page ranges extracted from the subject, or None if no page ranges were found.
    """
    page_ranges = None
    if "-p=" in subject:
        page_ranges = subject.split("-p=")[1].split()[0]
    return page_ranges

# Function to check for the page size


def get_page_size(subject):
    """
    Returns the page size specified in the email subject, or 'A4' if no page size is specified.

    Args:
        subject (str): The subject line of the email.

    Returns:
        str: The page size specified in the email subject, or 'A4' if no page size is specified.
    """
    page_size = "A4"
    if "-m=" in subject:
        page_size = subject.split("-m=")[1].split()[0]
    return page_size


def delete_emails(mail):
    """
    Deletes emails with a specific subject from the selected mailbox.

    Args:
        mail (IMAP4_SSL): An instance of the IMAP4_SSL class representing the mailbox.

    Returns:
        None
    """
    # Select the mailbox to work with
    mail.select('inbox/Druck')

    # Search for emails with the specified subject
    status, messages = mail.search(None, f'SUBJECT {PRINT_COMMAND}"')
    try:
        for message in messages[0].split():
            mail.store(message, "+FLAGS", "\\Deleted")

        # Permanently delete the messages marked for deletion
        mail.expunge()
    except:
        logging.info("No emails to delete.")
        return

# Function to check emails and trigger printing

# Function to check emails and process attachments for printing


def check_emails():
    """
    Connects to an email server and searches for new emails with a specific subject line. 
    If new emails are found, it processes the attachments and prints them using a specific printer.
    """
    # Setup logging
    logging.info("--------------------")
    logging.info("Connecting to email server...")

    # Connect to email server
    mail = imaplib.IMAP4_SSL(SMTP_HOST, SMTP_PORT)
    mail.login(SMTP_USER, SMTP_PASS)
    mail.select('inbox/Druck')

    # Search for specific emails
    logging.info("Searching for emails...")
    status, data = mail.search(None, (f'UNSEEN SUBJECT {PRINT_COMMAND}'))
    if status != 'OK':
        logging.error("Error searching emails.")
        return

    # If no new emails, return
    if not data or not data[0]:
        logging.info("No new emails to process.")
        return

    # Process new emails
    mail_ids = data[0].split()
    logging.info(f"Found {len(mail_ids)} emails")
    for mail_id in mail_ids:
        logging.info(f"Fetching email with ID: {mail_id}")
        typ, response = mail.fetch(mail_id, '(BODY[])')
        if typ != 'OK':
            logging.error("Error fetching mail.")
            continue

        # Process email attachments
        for response_part in response:
            if isinstance(response_part, tuple):
                email_data = response_part[1]
                email_message = email.message_from_bytes(email_data)
                logging.info(
                    f"Email subject: {email_message['subject']} from {email_message['from']}")

                # Determine print settings from email subject
                is_bw = is_black_and_white(email_message['subject'])
                is_duplex_mode = is_duplex(email_message['subject'])
                copies = get_copies(email_message['subject'])
                page_ranges = get_page_ranges(email_message['subject'])
                page_size = get_page_size(email_message['subject'])

                # Process PDF attachments
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()
                    if filename and filename.endswith('.pdf'):
                        logging.info(f"Found PDF attachment: {filename}")
                        file_data = part.get_payload(decode=True)
                        saved_file_path = save_file_to_local(
                            'attachments', filename, file_data)
                        print_pdf('kr-gha-3ogii.local', saved_file_path,
                                  'color' if is_bw else 'bw', 'true' if is_duplex_mode else 'false',
                                  page_size, copies, page_ranges
                                  )
                        delete_file('attachments', filename)

    # Delete emails
    delete_emails(mail)
    logging.info("Emails deleted.")

    # Close email connection
    mail.close()
    mail.logout()
    logging.info("Finished processing emails.")


# Main loop
if __name__ == "__main__":
    while True:
        check_emails()
        time.sleep(REFRESH_INTERVAL)
