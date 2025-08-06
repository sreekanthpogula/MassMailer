# MassMailer

MassMailer is a Internal tool designed to send bulk emails efficiently and reliably. It is ideal for newsletters, notifications, and other mass communication needs.
This project uses Python's `smtplib` for sending emails and supports customizable templates and pdf's attachments.


## Features

- Send emails to multiple recipients in bulk
- Customizable email templates using Jinja2
- Support for attachments (PDF, images, etc.)
- Logging and error handling
- Dry-run mode to test email sending without actually sending emails
- Environment variable configuration for sensitive data
- Support for HTML and plain text emails
- Email validation to ensure correct recipient addresses

## Installation

```bash
git clone https://github.com/sreekanthpogula/massmailer.git
cd massmailer
python -m venv venv
source .venv/bin/activate  # On Windows use ` Source venv\Scripts\activate`
pip install -r requirements.txt
python script.py
```

## Usage

1. Configure your SMTP settings in `.env` using the `.env.example` file as a reference.
2. Prepare your recipient list and email template using the provided sample files.
3. Place your email template in the `templates` directory and your recipient list in the `excel_files` directory.
4. Run the application using the command:
   ```bash
   streamlit run app.py
   ```
5. Access the web interface at `http://localhost:8501` to send emails.

## Sample Files
- **Email Template**: Place your email template in the `templates` directory. Use Jinja2 syntax for dynamic content.
- **KRA Template**: A sample KRA Excel file is provided in the `excel_files` directory. You can download it from the web interface.


## Feature Requests
If you have any feature requests or improvements, please open an issue on GitHub. We welcome contributions and suggestions to enhance the functionality of MassMailer.

## Requirements
- Python 3.10 or higher
- smtplib
- Jinja2
- python-dotenv
- email-validator
- logging
- requests
- streamlit (for the web interface)
- pandas (for handling Excel files)
- Flask (for the web server)

## Contributing

Pull requests are welcome. For major changes, open an issue first to discuss what you would like to change.

## License
This project is licensed under the MIT License.