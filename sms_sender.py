import csv
import time
import random
import sys
from ipaddress import IPv4Address, AddressValueError 
from typing import List, Dict, Optional

# Third-party libraries
try:
    from pyairmore.request import AirmoreSession
    from pyairmore.services.messaging import MessagingService
except ImportError:
    print("Error: Required libraries (pyairmore) are not installed.")
    print("Please run: pip install pyairmore")
    sys.exit(1)

# --- Configuration & Defaults ---

# Define the default values (can be overridden by user input in main)
DEFAULT_IP_ADDRESS = "192.168.1.68"
DEFAULT_PORT = 2333 

# Define the messages as a list of strings for clarity and easy modification
MESSAGES = [
    "बिधयालय लाई आवश्यक पर्ने कम्प्युटर, ल्यपटप, इन्टेरयाकटिब प्यानल बोर्ड, स्मार्ट टिभी, साइन्स ल्याब सामाग्री हरु एकै छातामुनी सर्बसुलभ रुपमा उपलब्ध छ ।",
    "सम्पर्क : मेगाटेक, मेनरोड, \n बिराटनगर \n ९८२०७५६२८०, ९८२०७५६२८२ \n हजुरकोको खुशी हाम्रो चाहना \n Facebook page: http://pili.app "
]

# Random delay range between messages to avoid overwhelming the server/phone
DELAY_RANGE_SECONDS = (3, 7)

# --- Data Structures and Utilities ---

class Recipient:
    """A data class to hold recipient information."""
    def __init__(self, name: str, phone_number: str, index: int):
        self.name = name
        self.phone_number = phone_number
        self.index = index

    def __repr__(self):
        return f"Recipient(Name='{self.name}', Phone='{self.phone_number}')"

def validate_and_parse_csv(file_path: str) -> Optional[List[Recipient]]:
    """
    Reads and validates a CSV file, returning a list of Recipient objects.
    Assumes the phone number is in the third column (index 2).
    """
    recipients: List[Recipient] = []
    
    try:
        # Use 'encoding=utf-8' explicitly, which is robust for non-Latin characters
        with open(file_path, 'r', encoding='utf-8') as file:
            csvreader = csv.reader(file)
            next(csvreader)  # Assuming the first row is a header and skipping it
            
            for i, row in enumerate(csvreader, start=2): # Start from row 2 (after header)
                if len(row) < 3:
                    print(f"Skipping row {i}: Too few columns.")
                    continue
                
                # Basic data cleaning and validation
                phone_number = row[2].strip()
                name = row[0].strip() if row[0] else f"Recipient {i-1}"
                
                # Simple check for number length
                if not phone_number.isdigit() or len(phone_number) < 7:
                    print(f"Warning: Invalid phone number '{phone_number}' in row {i}. Skipping.")
                    continue

                recipients.append(Recipient(name=name, phone_number=phone_number, index=i))
                
        return recipients
        
    except FileNotFoundError:
        print(f"Error: The file path '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while reading the CSV: {e}")
        return None

# --- Core Logic ---

def connect_to_airmore(ip_str: str, port: int) -> Optional[AirmoreSession]:
    """Establishes and authorizes an Airmore session."""
    print(f"\nAttempting to connect to Airmore at {ip_str}:{port}...")
    try:
        ip = IPv4Address(ip_str)
        session = AirmoreSession(ip, port)
        
        # NOTE: session.is_server_running performs a socket check and doesn't raise a pyairmore exception.
        if not session.is_server_running:
            print("Error: Airmore server is not running or unreachable on the specified device/IP.")
            return None
            
        print("Airmore is running. Requesting authorization on the phone...")
        was_accepted = session.request_authorization() 
        
        if was_accepted:
            print("✅ Authorization accepted! Connection successful.")
            return session
        else:
            print("❌ Authorization was DENIED or TIMED OUT on the phone.")
            return None
            
    except AddressValueError:
        print(f"Error: Invalid IP address format: {ip_str}")
        return None
    except Exception as e:
        print(f"An unexpected connection error occurred: {e}")
        return None

def send_bulk_messages(session: AirmoreSession, recipients: List[Recipient], messages: List[str]):
    """Sends a sequence of messages to a list of recipients with random delays."""
    
    if not recipients:
        print("No valid recipients to send messages to.")
        return

    print(f"\n--- Starting Bulk Message Send to {len(recipients)} recipients ---")
    
    # Initialize service once, outside the loop
    service = MessagingService(session)
    
    successful_sends = 0
    failed_sends = 0
    
    for recipient in recipients:
        print(f"\nSending to: {recipient.name} ({recipient.phone_number}) [Row {recipient.index}]")
        
        try:
            for i, body in enumerate(messages):
                print(f"  -> Sending Message {i+1}/{len(messages)}...")
                
                # service.send_message calls session.send(), which raises AuthorizationException
                service.send_message(recipient.phone_number, body) 
                
                # Sleep between individual messages within a recipient's sequence
                if i < len(messages) - 1: 
                    delay = random.randint(*DELAY_RANGE_SECONDS)
                    print(f"  -> Waiting {delay} seconds...")
                    time.sleep(delay)
            
            print("  ✅ All messages sent successfully to this recipient.")
            successful_sends += 1

        except Exception as e:
            print(f"  ❌ CRITICAL ERROR: Authorization revoked or expired on the device. Stopping all further sends.")
            failed_sends += (len(recipients) - successful_sends)
            break # Stop the loop on critical authorization failure
        except ServerUnreachableException as e:
            # Catching the specific ServerUnreachableException
            print(f"  ❌ CRITICAL ERROR: Server became unreachable during send. Stopping all further sends. Error: {e}")
            failed_sends += (len(recipients) - successful_sends)
            break # Stop the loop on critical failure
        except Exception as e:
            print(f"  ❌ ERROR: Failed to send message(s) to {recipient.phone_number}. Error: {e}")
            failed_sends += 1
            # Continue to the next recipient
    
    print("\n--- Bulk Send Summary ---")
    print(f"Total Recipients Attempted: {len(recipients)}")
    print(f"Successful Sends: {successful_sends}")
    print(f"Failed/Skipped Sends: {failed_sends}")
    print("-------------------------")

# --- Main Execution ---

def main():
    """The main function controlling the program flow, now with configurable connection details."""
    print("✨ Professional Airmore Bulk SMS Sender ✨")
    
    # 1. Get user input for configuration
    print(f"\n--- Configuration ---")
    
    ip_input = input(f"Enter Airmore IP Address (Default: {DEFAULT_IP_ADDRESS}): ").strip()
    airmore_ip = ip_input if ip_input else DEFAULT_IP_ADDRESS

    port_input = input(f"Enter Airmore Port (Default: {DEFAULT_PORT}): ").strip()
    airmore_port = int(port_input) if port_input.isdigit() else DEFAULT_PORT

    file_path = input("Enter the path to the CSV file (e.g., numbers.csv): ").strip()
    if not file_path:
        print("File path cannot be empty. Exiting.")
        return

    # 2. Parse Data
    recipients = validate_and_parse_csv(file_path)
    if not recipients:
        print("Could not process recipient data. Exiting.")
        return
    
    print(f"Loaded {len(recipients)} valid recipients from the CSV.")

    # 3. Connect
    session = connect_to_airmore(airmore_ip, airmore_port)
    if not session:
        print("Could not establish a successful Airmore session. Exiting.")
        return

    # 4. Send Messages
    send_bulk_messages(session, recipients, MESSAGES)
    
    print("\nProgram finished.")


if __name__ == "__main__":
    main()
