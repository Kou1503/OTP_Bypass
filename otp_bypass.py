import os
import imaplib
import email
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By

# Function to fetch OTP from Outlook
def fetch_otp_from_outlook():
    # Get email and password from environment variables
    email_user = os.environ.get('OUTLOOK_EMAIL')  # Fetch email from env variable
    email_pass = os.environ.get('OUTLOOK_PASSWORD')  # Fetch password from env variable

    # Connect to the Outlook IMAP server
    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(email_user, email_pass)
    
    mail.select("inbox")
    result, data = mail.search(None, 'UNSEEN')  # Search for unseen emails
    email_ids = data[0].split()
    
    if email_ids:
        latest_email_id = email_ids[-1]  # Get the latest email
        result, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Assuming the OTP is in the email subject or body
        subject = msg['subject']
        body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
        else:
            body = msg.get_payload(decode=True).decode()
        
        # Extract the OTP (Assuming it's a 6-digit number)
        otp = ''.join(filter(str.isdigit, body))  # Simple extraction
        return otp
    return None

# Function to automate pasting OTP into a website
def paste_otp(otp):
    # Copy OTP to clipboard
    pyperclip.copy(otp)

    # Set up the Opera GX WebDriver
    options = webdriver.ChromeOptions()
    options.binary_location = "C:/path/to/your/opera.exe"  # Update with the path to your Opera GX browser
    driver = webdriver.Opera(options=options)

    driver.get('https://example.com/login')  # Replace with your target URL

    # Wait for the page to load
    time.sleep(5)

    # Find the input field for OTP (Update the selector as needed)
    otp_input = driver.find_element(By.ID, 'otpInput')  # Update with the actual ID
    otp_input.click()

    # Paste the OTP
    otp_input.send_keys(pyperclip.paste())

# Main function
if __name__ == '__main__':
    while True:
        otp = fetch_otp_from_outlook()  # Fetch the OTP
        if otp:
            paste_otp(otp)  # Paste the OTP
        time.sleep(30)  # Check every 30 seconds or adjust as needed
