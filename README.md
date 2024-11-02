## mail_to_rm

This is a script to download attachments from emails from a given sender and save them to a given folder. It also converts the attachments to PDF if they are not already PDFs and uploads to my Remarkable tablet.

## Requirements

- **Runs only on Mac**: converts to PDF using applescripts
- Python
- Get your Gmail App password
- pyRMs GO dependency from https://github.com/ddvk/rmapi

## Usage

0. Install the necessary packages with `pip install -r requirements.txt`
0. Create a `.env` file with your email address and password
1. Run `main.py`