# emailer

Script to send a personalized email (with attachments) to several recipients.

Usage:
```bash
python3 emailer.py
```

## email data
The email configuration file.

Example:
```text
Source_addr: the-sending-email-address@example.com
Password: password
Subject: example-subject
Attachment_paths: first-file.pdf, second-file.pdf, ...
Salute: your-salute
Content: the-content-of-the-email
new-lines are possible!
```

## email list
Excel file with the name of the recipient in the first column, and the email address in the second column.

## TODO list:

* Make configuration file a JSON file.
* Speed up the email-sending process
