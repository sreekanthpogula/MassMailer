# MassMailer

MassMailer is a tool designed to send bulk emails efficiently and reliably. It is ideal for newsletters, notifications, and other mass communication needs.

## Features

- Send emails to multiple recipients
- Customizable email templates
- Support for attachments
- Logging and error handling

## Installation

```bash
git clone https://github.com/yourusername/massmailer.git
cd massmailer
npm install
```

## Usage

1. Configure your SMTP settings in `.env`.
2. Prepare your recipient list and email template.
3. Run the mailer:

```bash
npm start
```

## Configuration

Create a `.env` file with the following:

```
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USER=your_username
SMTP_PASS=your_password
```

## Contributing

Pull requests are welcome. For major changes, open an issue first to discuss what you would like to change.

## License

[MIT](LICENSE)