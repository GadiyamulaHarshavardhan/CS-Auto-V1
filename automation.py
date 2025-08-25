import os
import smtplib
import pandas as pd
import re
import time
import difflib
import logging
from unidecode import unidecode
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from datetime import datetime

# -----------------------
# Configuration
# -----------------------
load_dotenv()

# Validate required environment variables
EMAIL_ADDRESS = os.getenv("GMAIL_SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
if not EMAIL_ADDRESS:
    raise ValueError("GMAIL_SENDER_EMAIL not set in .env")
if not EMAIL_PASSWORD:
    raise ValueError("GMAIL_APP_PASSWORD not set in .env")

# Mode: Set TEST_MODE=false in .env to send real emails
TEST_MODE = os.getenv("TEST_MODE", "True").lower() == "true"

# Paths
cert_folder = "renamed_certificates"
excel_path = "carrer time machine hackathon attendence.xlsx"
log_file = f"email_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
match_report_file = f"match_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f"certificate_sender_{datetime.now().strftime('%Y%m%d')}.log"),
        logging.StreamHandler()
    ]
)

# -----------------------
# Utility Functions
# -----------------------

def normalize_name(name):
    """Normalize name: lowercase, ASCII, remove titles/suffixes, clean spaces"""
    if pd.isna(name):
        return ""
    name = str(name).strip()
    name = unidecode(name).lower()
    # Remove common titles and suffixes
    name = re.sub(r'\b(mr|mrs|ms|miss|dr|prof|sir|dame|jr|sr|ii|iii|iv|v|phd|md|dds|dvm)\b\.?\s*', '', name)
    # Keep only letters and spaces
    name = re.sub(r'[^a-z\s]', '', name)
    # Normalize whitespace
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def is_valid_email(email):
    """Validate email format"""
    if pd.isna(email):
        return False
    email = str(email).strip().lower()
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, email))

def clean_email(email):
    """Fix common email typos"""
    if pd.isna(email):
        return ""
    email = str(email).strip().lower()
    email = re.sub(r'\s+', '', email)
    corrections = {
        r'gmsil\.com': 'gmail.com',
        r'gnail\.com': 'gmail.com',
        r'@gmail,com': '@gmail.com',
        r'@yahoo,com': '@yahoo.com',
        r'@outlook,com': '@outlook.com',
        r'yahoo\s+com': 'yahoo.com',
        r'hotmail\s+com': 'hotmail.com',
        r'outlook\s+com': 'outlook.com'
    }
    for pattern, repl in corrections.items():
        email = re.sub(pattern, repl, email)
    return email if is_valid_email(email) else ""

def build_cert_map(folder):
    """
    Build map: normalized name → full filename
    Handles: Certificate_Harsha_pg001.pdf → harsha
    """
    certs = {}
    if not os.path.exists(folder):
        logging.error(f"Certificate folder '{folder}' not found!")
        return certs

    for filename in os.listdir(folder):
        if not filename.lower().endswith(".pdf"):
            continue

        base = os.path.splitext(filename)[0]
        # Extract name part: remove "Certificate_" and "_pgXXX"
        cleaned = re.sub(r'^certificate[_\-\s]*', '', base, flags=re.IGNORECASE)
        cleaned = re.sub(r'_pg\d+$', '', cleaned)
        cleaned = re.sub(r'_\d+$', '', cleaned)  # Also remove _123
        cleaned = cleaned.strip()

        norm_key = normalize_name(cleaned)
        if norm_key:
            certs[norm_key] = filename

    logging.info(f"📁 Certificate map built: {len(certs)} entries")
    return certs

def fuzzy_match(name, cert_keys, threshold=0.75):
    """Fuzzy match with higher threshold for precision"""
    matches = difflib.get_close_matches(name, cert_keys, n=1, cutoff=threshold)
    return matches[0] if matches else None

def extract_first_last_tokens(name):
    """Get first and last word for partial matching"""
    parts = name.split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[-1]}"
    return parts[0] if parts else ""

def validate_data_and_certificates(df, cert_map):
    """
    Match names from Excel to certificate files
    Column 0: Name
    Column 1: Email
    """
    results = []
    cert_keys = list(cert_map.keys())

    # Auto-detect header
    if len(df) > 0:
        first_row = df.iloc[0]
        if "name" in str(first_row.iloc[0]).lower() and "email" in str(first_row.iloc[1]).lower():
            start_idx = 1
            logging.info("✅ Header detected. Skipping first row.")
        else:
            start_idx = 0
    else:
        logging.error("❌ Excel file is empty!")
        return pd.DataFrame()

    for idx in range(start_idx, len(df)):
        row = df.iloc[idx]
        name_raw = str(row.iloc[0]).strip() if len(row) > 0 else ""
        email_raw = row.iloc[1] if len(row) > 1 else ""

        email = clean_email(email_raw)
        norm_name = normalize_name(name_raw)

        result = {
            'index': idx,
            'original_name': name_raw,
            'original_email': email_raw,
            'cleaned_email': email,
            'normalized_name': norm_name,
            'valid_email': bool(email),
            'certificate_found': False,
            'certificate_filename': None,
            'match_type': None,
            'match_confidence': 0.0,
            'status': 'pending'
        }

        if not result['valid_email']:
            result['status'] = 'invalid_email'
            results.append(result)
            continue

        # Try 1: Exact match
        if norm_name in cert_map:
            result['certificate_found'] = True
            result['certificate_filename'] = cert_map[norm_name]
            result['match_type'] = 'exact'
            result['match_confidence'] = 1.0
            result['status'] = 'matched'
        else:
            # Try 2: Fuzzy match
            fuzzy_key = fuzzy_match(norm_name, cert_keys, threshold=0.7)
            if fuzzy_key:
                confidence = difflib.SequenceMatcher(None, norm_name, fuzzy_key).ratio()
                result['certificate_found'] = True
                result['certificate_filename'] = cert_map[fuzzy_key]
                result['match_type'] = 'fuzzy'
                result['match_confidence'] = confidence
                result['status'] = 'matched' if confidence >= 0.7 else 'low_confidence_match'
            else:
                # Try 3: Partial match (first name only)
                first_name = norm_name.split()[0] if norm_name else ""
                partial_match = [k for k in cert_keys if k.startswith(first_name) or first_name in k]
                if partial_match:
                    best_match = max(partial_match,
                                   key=lambda k: difflib.SequenceMatcher(None, norm_name, k).ratio())
                    confidence = difflib.SequenceMatcher(None, norm_name, best_match).ratio()
                    if confidence >= 0.6:
                        result['certificate_found'] = True
                        result['certificate_filename'] = cert_map[best_match]
                        result['match_type'] = 'partial'
                        result['match_confidence'] = confidence
                        result['status'] = 'matched'
                    else:
                        result['status'] = 'certificate_not_found'
                else:
                    result['status'] = 'certificate_not_found'

        results.append(result)

    return pd.DataFrame(results)

# -----------------------
# Email Functions
# -----------------------

def send_certificate(name, email, cert_path, cert_filename, server):
    """Send professional HTML email with full-page background image"""
    subject = "Certificate of Participation – Career Time Machine Hackathon 2025"

    # Hosted background image (from ImgBB)
    bg_image_url = "https://i.ibb.co/bjy782Xj/hackathon-2-2.jpg"

    # HTML Email Body
    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        <title>Certificate Issued</title>
    </head>
    <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                margin: 0; 
                padding: 0; 
                background-image: url('{bg_image_url}'); 
                background-size: cover; 
                background-position: center; 
                background-repeat: no-repeat; 
                color: white; 
                line-height: 1.6;
                min-height: 100vh;">
        
        <!-- Dark overlay for readability -->
        <div style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                    background: rgba(0, 0, 0, 0.7); 
                    z-index: 1;"></div>

        <!-- Main Content -->
        <div style="position: relative; z-index: 2; 
                    max-width: 650px; 
                    margin: 40px auto; 
                    padding: 30px; 
                    color: white; 
                    text-align: center;">
            
           

            <!-- Main Title -->
            <h1 style="margin: 10px 0; 
                       font-size: 28px; 
                       font-weight: 600; 
                       color: #ffffff;">
                Certificate of Participation
            </h1>
            
            <!-- Event Name -->
            <p style="margin: 5px 0; 
                      font-size: 18px; 
                      color: #ffffff; 
                      font-weight: 500;">
                Career Time Machine Hackathon 2025
            </p>
            
            <!-- Organized By -->
            <p style="margin: 15px 0 25px; 
                      font-size: 14px; 
                      color: rgba(255, 255, 255, 0.9);">
                Organized by <strong>Darion Playshops</strong><br>
                in collaboration with <strong>Hindu College of Engineering & Technology</strong>
            </p>

            <!-- Message -->
            <div style="background: rgba(0, 0, 0, 0.5); 
                        border-radius: 10px; 
                        padding: 20px; 
                        text-align: left; 
                        display: inline-block; 
                        max-width: 100%;">
                <p style="font-size: 16px; margin: 0 0 15px; color: #ffffff;">
                    Hello <strong>{name}</strong>,
                </p>

                <p style="font-size: 16px; margin: 0 0 15px; color: #ffffff;">
                    Congratulations on participating in the <strong>Career Time Machine Hackathon 2025</strong>, organized by <strong>Darion Playshops</strong> in collaboration with <strong>Hindu College of Engineering and Technology</strong> on <strong>23 August 2025</strong>!
                </p>

                <p style="font-size: 16px; margin: 0 0 15px; color: #ffffff;">
                    Your creativity, teamwork, and curiosity in exploring future careers and innovations through time-travel inspired AI solutions were truly inspiring. Whether you reached the initial stages or advanced to the final levels, you have taken a powerful step toward building the future of technology.
                </p>

                <p style="font-size: 16px; margin: 0 0 15px; color: #ffffff;">
                    Please find your <strong>Certificate of Participation</strong> attached to this email. May it remind you that you are capable of creating meaningful, future-driven solutions that make a difference.
                </p>

                <p style="font-size: 16px; margin: 0 0 15px; color: #ffffff;">
                    For any doubts, queries, or feedback, feel free to reach us at <a href="mailto:workshop.darion@gmail.com" style="color: #00c8ff;">workshop.darion@gmail.com</a>.
                </p>

                <div style="text-align: center; margin: 20px 0;">
                    <span style="display: inline-block; 
                                 background: #0d3b66; 
                                 color: white; 
                                 padding: 10px 20px; 
                                 border-radius: 6px; 
                                 font-weight: bold; 
                                 font-size: 15px;">
                        Never stop building. Never stop imagining.
                    </span>
                </div>

                <p style="font-size: 16px; color: #ffffff;">
                    The world needs what you're capable of.
                </p>
            </div>

            <!-- Footer with Logos -->
<div style="position: relative; z-index: 2; 
            margin-top: 40px; 
            padding: 20px; 
            text-align: center; 
            font-size: 14px; 
            color: rgba(255, 255, 255, 0.9);">
    
    <!-- Logos -->
    <div style="display: flex; 
                justify-content: center; 
                align-items: center; 
                gap: 40px; 
                margin-bottom: 15px;">
        
        <!-- Hindu College Logo -->
        <img src="https://i.ibb.co/DDXMbhk5/hackathon-2-3.png" 
             alt="Hindu College of Engineering & Technology" 
             style="height: 82px;" />

    
    </div>

    <!-- Closing Message -->
    <p style="margin: 0 0 10px; font-size: 16px;">
        With pride and appreciation,<br>
        <strong style="color: #00c8ff;">Team Darion</strong>
    </p>

    <!-- Links -->
    <p style="margin: 0;">
        <a href="https://darion.in" target="_blank" 
           style="color: #00c8ff; text-decoration: none; margin: 0 10px; font-weight: 500;">
           darion.in
        </a> | 
        <a href="https://hinduengg.com" target="_blank" 
           style="color: #00c8ff; text-decoration: none; margin: 0 10px; font-weight: 500;">
           hinduengg.com
        </a>
    </p>
</div>
        </div>
    </body>
    </html>
    """

    # Plain text fallback
    plain_body = f"""Hello {name},

Congratulations on participating in the Career Time Machine Hackathon 2025, organized by Darion Playshops in collaboration with Hindu College of Engineering and Technology on 23 August 2025!

Your creativity, teamwork, and curiosity in exploring future careers and innovations through time-travel inspired AI solutions were truly inspiring. Whether you reached the initial stages or advanced to the final levels, you have taken a powerful step toward building the future of technology.

Please find your Certificate of Participation attached to this email. May it remind you that you are capable of creating meaningful, future-driven solutions that make a difference.

For any doubts, queries, or feedback, feel free to reach us at workshop.darion@gmail.com.

Never stop building. Never stop imagining. The world needs what you're capable of.

With pride and appreciation,
Team Darion
darion.in | hinduengg.com
"""

    try:
        msg = MIMEMultipart("alternative")
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = email
        msg['Subject'] = subject

        # Attach both plain and HTML versions
        part1 = MIMEText(plain_body, "plain")
        part2 = MIMEText(html_body, "html")
        msg.attach(part1)
        msg.attach(part2)

        # Attach certificate
        with open(cert_path, "rb") as attachment:
            cert_part = MIMEBase('application', 'octet-stream')
            cert_part.set_payload(attachment.read())
            encoders.encode_base64(cert_part)
            cert_part.add_header(
                'Content-Disposition',
                f'attachment; filename="{cert_filename}"'
            )
            msg.attach(cert_part)

        # Send or simulate
        if TEST_MODE:
            logging.info(f"[TEST MODE] Would send to {email} | Certificate: {cert_filename}")
            return True
        else:
            server.send_message(msg)
            logging.info(f"✅ Sent to {email} | Attached: {cert_filename}")
            return True

    except Exception as e:
        logging.error(f"❌ Failed to send to {email}: {e}")
        return False

# -----------------------
# MAIN SCRIPT
# -----------------------

def main():
    logging.info("🚀 Starting certificate email process")

    # Load Excel
    try:
        df = pd.read_excel(excel_path, header=None)
        logging.info(f"📊 Loaded Excel with {len(df)} rows and {len(df.columns)} columns")
    except Exception as e:
        logging.error(f"❌ Failed to load Excel: {e}")
        return

    if len(df.columns) < 2:
        logging.error("❌ Excel must have at least 2 columns: Name and Email")
        return

    # Load certificates
    try:
        available_certs = build_cert_map(cert_folder)
        if not available_certs:
            logging.error("❌ No valid certificate files found!")
            return
        logging.info(f"📁 Found {len(available_certs)} certificate(s)")
        logging.info(f"📄 Sample: {list(available_certs.values())[:3]}")
    except Exception as e:
        logging.error(f"❌ Failed to load certificates: {e}")
        return

    # Validate and match
    validation_df = validate_data_and_certificates(df, available_certs)
    validation_df.to_csv(match_report_file, index=False)
    logging.info(f"📋 Match report saved to {match_report_file}")

    # Summary
    status_counts = validation_df['status'].value_counts()
    logging.info("📋 Match Summary:")
    for status, count in status_counts.items():
        logging.info(f"  {status}: {count}")

    # Warnings
    missing = len(validation_df[validation_df['status'] == 'certificate_not_found'])
    low_conf = len(validation_df[validation_df['status'] == 'low_confidence_match'])
    invalid = len(validation_df[validation_df['status'] == 'invalid_email'])

    if missing or low_conf or invalid:
        logging.warning(f"⚠️ Issues: {missing} missing, {low_conf} low-confidence, {invalid} invalid emails")
        if not TEST_MODE:
            response = input("⚠️  High failure count. Continue? (yes/no): ")
            if response.lower() != 'yes':
                logging.info("🛑 Process cancelled by user.")
                return

    # Connect to SMTP
    server = None
    if not TEST_MODE:
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            logging.info("✅ SMTP login successful")
        except Exception as e:
            logging.error(f"❌ SMTP login failed: {e}")
            return
    else:
        logging.info("🧪 TEST MODE: No emails will be sent")

    # Send emails
    successful = failed = skipped = 0
    with open(log_file, "a", encoding="utf-8") as log:
        log.write(f"\n\n=== Certificate Email Process - {datetime.now()} ===\n")
        log.write("Status | Email | Name | Details\n")
        log.write("-" * 80 + "\n")

        for _, row in validation_df.iterrows():
            if row['status'] not in ['matched', 'low_confidence_match']:
                skipped += 1
                log.write(f"SKIPPED | {row['cleaned_email']} | {row['original_name']} | {row['status']}\n")
                continue

            cert_filename = row['certificate_filename']
            cert_path = os.path.join(cert_folder, cert_filename)

            if not os.path.exists(cert_path):
                logging.error(f"❌ Certificate not found: {cert_path}")
                failed += 1
                log.write(f"FAILED | {row['cleaned_email']} | {row['original_name']} | File missing\n")
                continue

            try:
                success = send_certificate(
                    name=row['original_name'],
                    email=row['cleaned_email'],
                    cert_path=cert_path,
                    cert_filename=cert_filename,
                    server=server
                )
                if success:
                    successful += 1
                    log.write(f"SUCCESS | {row['cleaned_email']} | {row['original_name']} | {cert_filename}\n")
                else:
                    failed += 1
                    log.write(f"FAILED | {row['cleaned_email']} | {row['original_name']} | Send failed\n")
            except Exception as e:
                failed += 1
                log.write(f"FAILED | {row['cleaned_email']} | {row['original_name']} | Exception: {str(e)}\n")
                logging.error(f"❌ Exception for {row['cleaned_email']}: {e}")

            time.sleep(1.5)  # Rate limiting

    # Close server
    if server:
        server.quit()
        logging.info("📤 SMTP connection closed")

    # Final summary
    logging.info("=" * 60)
    logging.info("✅ EMAIL PROCESSING COMPLETED")
    logging.info(f"📬 Successful: {successful}")
    logging.info(f"❌ Failed: {failed}")
    logging.info(f"🚫 Skipped: {skipped}")
    logging.info(f"📊 Total: {successful + failed + skipped}")
    logging.info(f"📄 Log: {log_file}")
    logging.info(f"📋 Report: {match_report_file}")
    logging.info("🎉 All done!")

if __name__ == "__main__":
    main()