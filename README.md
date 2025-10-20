# Professional Data Processing and Messaging Tools

This repository contains three professional-grade Python scripts designed to automate common business tasks efficiently:

- **Bulk SMS Sender**: Sends SMS messages via an Android phone using the Airmore application.
- **Bulk Email Sender**: Sends bulk personalized emails with attachments via an SMTP server (configured for Gmail).
- **Excel-to-CSV Converter**: Extracts specified columns from an Excel file and exports the data to a new CSV file.

---

## ⚙️ Prerequisites

- Python 3.8 or higher installed on your system.

---

## Installation

Install all required dependencies using pip and the provided `requirements.txt` file:


This will install the following packages:

- `pyairmore`
- `openpyxl`

---

## 🚀 1. Bulk SMS Sender (`sms_sender.py`)

This script uses the `pyairmore` library to send SMS messages via an Android phone connected through the Airmore app.

### Prerequisites

- **Airmore App**: Install and run the Airmore app on your Android device.
- **Network**: Ensure your computer and Android phone are connected to the same Wi-Fi network.
- **IP Address**: Find your phone's local IP address (e.g., 192.168.1.68) and configure it in the script.

### Usage

- **Configuration**:
  - Edit the `AIRMORE_IP_ADDRESS` constant in the script to match your phone’s local IP.
  - Customize the `MESSAGES` list with your message content.
  - Adjust `DELAY_RANGE_SECONDS` to set the wait time between sending messages.
- **CSV Format**:
  - Your recipient CSV must have phone numbers in the third column (index 2).
- **Run the script**:


The script will prompt for the CSV file path and then ask for authorization on your phone.

---

## 📧 2. Bulk Email Sender (`email_sender.py`)

This script sends bulk personalized emails using Gmail’s SMTP server.

### Security and Authentication Setup

- You MUST generate a Google App Password, not use your Gmail account password.
- Steps:
  1. Go to your Google Account Security Settings.
  2. Enable 2-Step Verification if it is not already enabled.
  3. Generate an App Password for the "Mail" app on a "Other" device.
- **Important**: Do not hardcode your app password in production. For demonstration, a placeholder is used as `HARDCODED_APP_PASSWORD` in the script.

### Usage

- **Configuration**:
  - Customize `DEFAULT_SUBJECT` and `DEFAULT_BODY_TEXT` inside the script.
  - Adjust `DELAY_RANGE_SECONDS` to control the sending pace (recommended to keep it high to prevent spam detection).
  - Update `HARDCODED_APP_PASSWORD` with your generated app password.
- **CSV Format**:
  - Recipient's name should be in the first column (index 0).
  - Recipient's email should be in the second column (index 1).
- **Run the script**:


The script will prompt for:
- Your sender email address.
- The recipient CSV file path.
- (Optional) Attachment file path.

---

## 📊 3. Excel-to-CSV Converter (`excel_to_csv.py`)

This utility extracts selected columns from an Excel sheet and saves the output as a CSV file.

### Usage

Run the script:


You will be guided through these prompts:

- **Excel File Path**: Enter the full path to your `.xlsx` file.
- **Sheet Selection**: Specify the sheet by name (e.g., `"Data"`) or by 0-based index (e.g., `"0"`).
- **Column Selection**: Specify columns using letters (A, B, C) or numbers (1, 2, 3), separated by commas (e.g., `A,C,F` or `1,3,6`).
- **Output File Name**: Enter the name for the exported `.csv` file.

### Error Handling

The script includes robust error checks for:

- File not found.
- Invalid Excel file format.
- Invalid sheet name/index or column selection.

---

Feel free to explore and customize these scripts to fit your business automation needs.
