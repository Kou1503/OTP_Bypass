import os
import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import pyautogui

# Clear old environment variables
os.environ.pop("USERNAME", None)
os.environ.pop("PASSWORD", None)

# Load environment variables
load_dotenv()
LOGIN_URL = os.getenv("LOGIN_URL")
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
OTP_WAIT_TIME = int(os.getenv("OTP_WAIT_TIME", 10))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", 5))

# Debugging
print(f"LOGIN_URL: {LOGIN_URL}")
print(f"USERNAME: {USERNAME}")
print(f"PASSWORD: {PASSWORD}")

# Initialize Chrome WebDriver
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)  # Keeps the browser open

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Open the login page
driver.get(LOGIN_URL)

# Wait for username field and enter username
username_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.NAME, "userid"))
)
username_field.send_keys(USERNAME)

# Wait for password field and enter password
password_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.NAME, "pwd"))
)
password_field.send_keys(PASSWORD)

# Click the Sign In button
sign_in_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ps_submit_button"))
)
sign_in_button.click()

# Wait for OTP input field
otp_field = WebDriverWait(driver, OTP_WAIT_TIME).until(
    EC.presence_of_element_located((By.NAME, "otp"))
)

# Switch to Outlook and get the OTP
time.sleep(5)  # Give time to switch manually if needed
pyautogui.hotkey("alt", "tab")  # Switch to Outlook

time.sleep(2)  # Wait for Outlook to be active
pyautogui.hotkey("ctrl", "e")  # Focus search bar
pyautogui.write("Your OTP Code")  # Search for the OTP email
pyautogui.press("enter")

time.sleep(2)  # Wait for search results
pyautogui.press("down")  # Select the first email
pyautogui.press("enter")

time.sleep(2)  # Wait for email to open
pyautogui.hotkey("ctrl", "a")  # Select all text
pyautogui.hotkey("ctrl", "c")  # Copy text

time.sleep(1)
otp_text = pyautogui.paste()  # Get copied text
otp_code = re.search(r'\b\d{6}\b', otp_text)  # Extract 6-digit OTP

if otp_code:
    otp_code = otp_code.group()
    print(f"Extracted OTP: {otp_code}")
    otp_field.send_keys(otp_code)

    # Click Submit or Continue (adjust selector if needed)
    submit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Submit") or contains(text(),"Continue")]'))
    )
    submit_button.click()
else:
    print("Failed to extract OTP")

# Keep the browser open for debugging
input("Press Enter to close the browser...")
driver.quit()
