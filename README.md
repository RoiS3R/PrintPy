# Email to Print Automation Script

This Python script automatically checks an email inbox for new messages with a specific subject line, downloads PDF attachments from these emails, and sends them to a specified printer.

### Prerequisites

Before you start using this script, ensure you have the following prerequisites:

* Python 3.x installed on your system.
* Access to an SMTP server (like Gmail, iCloud, etc.) with the ability to use IMAP.
* A printer set up on your network that can be accessed via CUPS (Common Unix Printing System).
* The python-cups and python-dotenv packages installed in Python.

### Installation

#### 1. Clone the Repository:
Clone this repository to your local machine or download the script file.
#### 2. Install Required Packages:
Run pip install python-cups python-dotenv to install the necessary Python packages.
#### 3. Environment Variables:
Create a .env file in the same directory as the script.
Add the following environment variables with your SMTP details:
```bash
makefile
Copy code
SMTP_USER=your_email@example.com
SMTP_PASS=your_password
SMTP_HOST=smtp.example.com
SMTP_PORT=993
SMTP_TLS=true

# printer uri and name
PRINTER1_URI=ipp://192.168.1.100:631/printers/HP_LaserJet_Pro_MFP_M428fdw # for example
PRINTER1_NAME=HP LaserJet Pro
# PRINTER2_URI=
# PRINTER2_NAME=
# PRINTER3_URI=
# PRINTER3_NAME=
```

#### 4. Printer Configuration:
Replace the placeholder values in the script with your printer's details, including its URI and name.

### Usage

#### 1. Run the Script:
Execute the script with the command python script_name.py.
The script will continuously check for new emails every 10 seconds.
#### 2. Sending Print Jobs:
* Send an email to the configured SMTP_USER email address.
* Use the subject line `**Print` to trigger the print function.
* The following options are (optionally) available by including them in the subject line:
```bash
-c=color or -c=bw #to select either colour or black and white print (default: bw)
-d=duplex or -d=simplex #to select either duplex or simplex print (default: simplex)
-m=a4 , -m=a3 or -m=a5 #to select the page size used for printing (default: a4)
-copies=n #to select the number of copies to print (default: 1)
-p=1-4 or -page_range=1,3-5 #to select the pages to print add either individual numbers (1), ranges (1-4) or a combination (1,4-5). Caution: Do **not** use spaces.
```
* Example: the E-Mail subject `**Print -c=bw -copies=4` prints _4 copies_ of the _complete_ document in _black and white_ on _A4_, _simplex_
* Attach one or more PDF document to the email. Be aware: The options are applied to all attached documents, so send them in seperate E-Mails when having specific requirements
#### 3. Automatic Printing:
The script will download the PDF attachment and send it to the specified printer.
#### 4. Stop the Script:
To stop the script, use the keyboard interrupt Ctrl + C in the terminal.

### Notes

* Ensure that the email account used has IMAP access enabled.
* For security reasons, it's recommended to use an app-specific password if your email provider supports it.
* The script currently handles only PDF attachments and prints the entire document.
* Modify the script to suit your specific needs, like handling different file types or print settings.

### License