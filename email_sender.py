import csv
import time
import random
import sys
import os
import smtplib
import ssl
from typing import List, Optional

# Email handling libraries
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- Configuration & Constants ---

# SMTP Settings
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465 # Standard port for SSL/TLS

# Default Email Content
DEFAULT_SUBJECT = 'Important Update from [Your Company]'
DEFAULT_BODY_TEXT = """
Dear Recipient,

This is a test email sent from our automated professional script.
We appreciate your time.

Best regards,
Automated Sender
"""

# Security Note: NEVER hardcode your real password in a professional script.
# The following is for demonstration of where a dedicated 'App Password' goes.
# IMPORTANT: This password is a Gmail 'App Password', not your primary password.
HARDCODED_APP_PASSWORD = 'tvbprmvgogbjwgmj' # KEEP THIS OUTSIDE THE CODE IN A REAL APP!

# Delay range between sending emails to avoid being flagged as spam
DELAY_RANGE_SECONDS = (5, 15) # Increased delay for better server friendliness

# --- Data Structures and Utilities ---

class Recipient:
    """A data class to hold recipient information."""
    def __init__(self, name: str, email: str, index: int):
        self.name = name
        self.email = email
        self.index = index

    def __repr__(self):
        return f"Recipient(Name='{self.name}', Email='{self.email}')"

def validate_and_parse_csv(file_path: str) -> Optional[List[Recipient]]:
    """
    Reads and validates a CSV file, returning a list of Recipient objects.
    Assumes email address is in the second column (index 1).
    Assumes name is in the first column (index 0).
    """
    recipients: List[Recipient] = []

    try:
        # Using 'utf-8' and setting a default name if missing
        with open(file_path, 'r', encoding='utf-8') as file:
            csvreader = csv.reader(file)
            # Assuming the first row is a header and skipping it
            try:
                next(csvreader)
            except StopIteration:
                print("Error: CSV file is empty.")
                return None

            for i, row in enumerate(csvreader, start=2): # Start from row 2 (after header)
                if len(row) < 2:
                    print(f"Skipping row {i}: Too few columns. (Requires Name and Email)")
                    continue

                # Basic data cleaning and validation
                name = row[0].strip() if row[0].strip() else f"Recipient {i-1}"
                email_address = row[1].strip()

                # Simple check for a valid-looking email (contains '@' and '.')
                if '@' not in email_address or '.' not in email_address:
                    print(f"Warning: Invalid email '{email_address}' in row {i}. Skipping.")
                    continue

                recipients.append(Recipient(name=name, email=email_address, index=i))

        return recipients

    except FileNotFoundError:
        print(f"Error: The file path '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while reading the CSV: {e}")
        return None

def create_email_message(sender_email: str, recipient: Recipient, subject: str, body: str, attachment_path: Optional[str]) -> MIMEMultipart:
    """Constructs the MIMEMultipart object for sending."""
    em = MIMEMultipart()
    em['From'] = sender_email
    em['To'] = recipient.email
    em['Subject'] = subject

    # Attach body content
    # The body can be customized to greet the recipient by name (e.g., 'Dear {recipient.name}')
    em.attach(MIMEText(body.replace("Dear Recipient", f"Dear {recipient.name}"), 'plain'))

    # Handle Attachment
    if attachment_path and os.path.exists(attachment_path):
        try:
            with open(attachment_path, 'rb') as attach_file: # Open file as binary
                payload = MIMEBase('application', 'octet-stream')
                payload.set_payload(attach_file.read())
                encoders.encode_base64(payload) # Encode the attachment

                # Add payload header with original filename
                filename = os.path.basename(attachment_path)
                payload.add_header('Content-Disposition', 'attachment', filename=filename)
                em.attach(payload)
            print(f"  -> Attached file: {filename}")
        except Exception as e:
            print(f"  ❌ WARNING: Could not attach file '{attachment_path}'. Error: {e}")

    return em

# --- Core Logic ---

def send_bulk_emails(sender_email: str, app_password: str, recipients: List[Recipient], attachment_path: Optional[str]):
    """Sends a sequence of emails with a random delay."""

    if not recipients:
        print("No valid recipients to send emails to.")
        return

    print(f"\n--- Starting Bulk Email Send to {len(recipients)} recipients ---")

    # Create SSL context once
    context = ssl.create_default_context()
    successful_sends = 0
    failed_sends = 0

    try:
        # Establish connection outside the loop for efficiency
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
            print(f"Connecting to {SMTP_SERVER}...")
            smtp.login(sender_email, app_password)
            print("✅ Login successful. Starting send loop...")

            for recipient in recipients:
                print(f"\nSending to: {recipient.name} ({recipient.email}) [Row {recipient.index}]")

                try:
                    # Create the full email message
                    msg = create_email_message(
                        sender_email,
                        recipient,
                        DEFAULT_SUBJECT,
                        DEFAULT_BODY_TEXT,
                        attachment_path
                    )

                    # Send the email
                    smtp.sendmail(sender_email, recipient.email, msg.as_string())
                    print("  ✅ Email sent successfully.")
                    successful_sends += 1

                except Exception as e:
                    print(f"  ❌ ERROR: Failed to send email to {recipient.email}. Error: {e}")
                    failed_sends += 1
                    # Continue to the next recipient

                # Apply delay after each successful or failed attempt
                delay = random.randint(*DELAY_RANGE_SECONDS)
                print(f"  -> Waiting {delay} seconds...")
                time.sleep(delay)

    except smtplib.SMTPAuthenticationError:
        print("\n❌ CRITICAL ERROR: Authentication failed. Check your email address and App Password.")
        print("    (Note: You must use an App Password for Gmail, not your main account password.)")
    except smtplib.SMTPConnectError:
        print(f"\n❌ CRITICAL ERROR: Could not connect to {SMTP_SERVER}. Check network connection or SMTP settings.")
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: An unexpected error occurred during connection: {e}")

    print("\n--- Bulk Send Summary ---")
    print(f"Total Recipients Attempted: {len(recipients)}")
    print(f"Successful Sends: {successful_sends}")
    print(f"Failed Sends: {failed_sends}")
    print("-------------------------")


# --- Main Execution ---

def main():
    """The main function controlling the program flow."""
    print("✨ Professional Bulk Email Sender (SMTP) ✨")
    print("--- Configuration ---")
    print(f"SMTP Server: {SMTP_SERVER}")
    print(f"Delay Range: {DELAY_RANGE_SECONDS[0]}-{DELAY_RANGE_SECONDS[1]} seconds")
    print("---------------------")

    # Get user input
    file_path = input("Enter the path to the CSV file (e.g., list.csv): ").strip()
    sender_email = input("Enter your sender Email (e.g., your@gmail.com): ").strip()

    # NOTE: In a professional application, the password would be read from an
    # environment variable, configuration file, or secure vault, NOT hardcoded.
    # We'll use the hardcoded placeholder for functionality.
    app_password = HARDCODED_APP_PASSWORD

    # Get attachment path (Optional)
    attachment_path = input("Enter the full path to an attachment file (optional, leave blank to skip): ").strip() or None

    if not file_path or not sender_email:
        print("Required fields cannot be empty. Exiting.")
        return

    # 1. Parse Data
    recipients = validate_and_parse_csv(file_path)
    if not recipients:
        print("Could not process recipient data. Exiting.")
        return

    print(f"Loaded {len(recipients)} valid recipients from the CSV.")

    # 2. Send Messages
    send_bulk_emails(sender_email, app_password, recipients, attachment_path)

    print("\nProgram finished.")


if __name__ == "__main__":
    main()