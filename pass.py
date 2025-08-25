# test_smtp.py
import smtplib
from dotenv import load_dotenv
import os

load_dotenv()

EMAIL = os.getenv("GMAIL_SENDER_EMAIL")
PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

try:
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL, PASSWORD)
    print("✅ Login successful!")
    server.quit()
except Exception as e:
    print(f"❌ Login failed: {e}")