import os
import time
import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv

# Load environment variables from the config.env file
load_dotenv('config.env')

# Get credentials and configurations from environment variables
email = os.getenv('EMAIL')
password = os.getenv('PASSWORD')
website_url = os.getenv('WEBSITE_URL')
opera_path = os.getenv('OPERA_PATH')

# Function to get OTP from Outlook
def get_otp():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox
    messages = inbox.Items

    for message in messages:
        if "OTP" in message.Subject:  # Change this to match the subject of your OTP emails
            return message.Body  # Get the body of the email containing the OTP

    return None  # Return None if no OTP found

# Function to open the website and enter the OTP
def enter_otp(otp):
    # Set up Selenium to use Opera GX
    options = Options()
    options.binary_location = opera_path
    driver = webdriver.Opera(service=Service(), options=options)

    try:
        driver.get(website_url)  # Open the target website
        time.sleep(5)  # Wait for the page to load (adjust as needed)

        # Locate the OTP input field and enter the OTP
        otp_input = driver.find_element("name", "otp")  # Adjust based on the actual input field's name or identifier
        otp_input.send_keys(otp)
        
        submit_button = driver.find_element("id", "submit")  # Adjust based on the actual button's identifier
        submit_button.click()  # Click the submit button
        
    finally:
        time.sleep(5)  # Wait to see the result (adjust as needed)
        driver.quit()  # Close the browser

# Main function to run the program
if __name__ == "__main__":
    otp = get_otp()  # Fetch the OTP from Outlook
    if otp:
        print(f"OTP retrieved: {otp}")
        enter_otp(otp)  # Enter the OTP into the website
    else:
        print("No OTP found in inbox.")
