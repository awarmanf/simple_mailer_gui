# SimpleMailer

SimpleMailer is a lightweight email-sending application built with PyQt5 and QFluentWidgets, designed for sending emails through SMTP servers. This application supports multiple security types such as SSL, TLS, and plain text authentication. It also supports email attachments, making it suitable for personal or professional use.

## Features

- **Customizable SMTP Configuration:** Users can input SMTP host, security type, and credentials to send emails.
- **Multiple Security Types:** Supports SMTP over SSL, TLS, and plain text connections.
- **Attachment Support:** Attach files to emails with ease.
- **Modern UI:** Uses QFluentWidgets for a clean, fluent design experience.
- **Frameless Window:** Provides a modern, borderless window with custom close and minimize buttons.

## Screenshots

*(Include screenshots of the UI for better visual representation.)*

## Requirements

- Python 3.7 or later
- PyQt5
- qfluentwidgets
- qframelesswindow

You can install all the required packages by running:

```bash
pip install -r requirements.txt
```

## Installation

1. Clone this repository and Extract it

2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

3. Run the application:

```bash
python SimpleMailerApp.py
```

## Usage

1. Launch the application and enter the required details:
   - Full name
   - Email account
   - Password
   - SMTP host (e.g., `smtp.gmail.com`)
   - Security type (SSL, TLS, or plain text)
   - Recipients (separate multiple emails with commas)
   - Email subject
   - Email body (can include plain text or HTML)
   - Optional: Attach a file using the "Select File" button.

2. Click the "Send" button to send the email.

## Configuration

The application stores the following information in a `config.json` file, which is generated upon the first launch:

```json
{
  "full_name": "Your Name",
  "account_data": {
    "1": {
      "account_mail": "your-email@example.com",
      "account_password": "your-password"
    }
  },
  "SMTP_host": "smtp.example.com",
  "security_type": "SSL"
}
```

You can manually edit this file to update your settings or change the email account used to send emails.

## Security Notice

Make sure to handle sensitive information such as email passwords with care. Avoid hardcoding credentials or sharing configuration files with sensitive information. Consider using environment variables or other secure methods for credential management in production use.

## Known Issues

- Error handling during SMTP connection or email sending might not always provide detailed messages.
- The application only supports plain text emails; HTML emails are not yet supported.
- The user interface is optimized for a specific resolution. On smaller screens, some elements may appear misaligned.

## Future Plans

- Add support for HTML emails.
- Improve error handling for better user feedback.
- Enable resizing the application for smaller screens.
- Add more theme options (light/dark).