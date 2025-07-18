# MassMailer

MassMailer is a Internal tool designed to send bulk emails efficiently and reliably. It is ideal for newsletters, notifications, and other mass communication needs.
This project uses Python's `smtplib` for sending emails and supports customizable templates and pdf's attachments.


## Features

- Send emails to multiple recipients in bulk
- Customizable email templates using Jinja2
- Support for attachments (PDFs)
- Logging and error handling
- Dry-run mode to test email sending without actually sending emails
- Environment variable configuration for sensitive data
- Command-line interface for easy usage
- Support for HTML and plain text emails
- Email validation to ensure correct recipient addresses

## Installation

```bash
git clone https://github.com/yourusername/massmailer.git
cd massmailer
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
pip install -r requirements.txt
python script.py
```

## Usage

1. Configure your SMTP settings in `.env`.
2. Prepare your recipient list and email template.
3. Run the mailer:

```bash
python script.py --dry-run
python script.py
```

## Configuration

Create a `.env` file with the following:

```
MAIL_SERVER=mail_server_here
MAIL_PORT= mail_port_here
MAIL_USE_TLS= tls setting_here
MAIL_USERNAME= your email_here
MAIL_PASSWORD= your_email_password_here
MAIL_DEFAULT_SENDER_NAME= default sender name
MAIL_DEFAULT_SENDER_EMAIL= your sender email_here
MAIL_DEFAULT_RECEIVER_NAME= default receiver name
MAIL_DEFAULT_RECEIVER_EMAIL= your receiver email_here
```

## Feature Requests
If you have any feature requests or improvements, please open an issue on GitHub. We welcome contributions and suggestions to enhance the functionality of MassMailer.

## Requirements
- Python 3.10 or higher
- smtplib
- Jinja2
- python-dotenv
- email-validator
- argparse
- logging
- requests

## Contributing

Pull requests are welcome. For major changes, open an issue first to discuss what you would like to change.

## License

[MIT](LICENSE)