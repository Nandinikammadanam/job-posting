import os
import json
import base64
import re
import time
import pytz
from datetime import datetime
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import qn
from docx.oxml import parse_xml
import requests
from groq import Groq 
from typing import List, Dict, Optional
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Load environment variables
load_dotenv()


# Define multiple credential sets - DO NOT CHANGE USERNAME AND PASSWORD
CREDENTIAL_SETS = [
    {
        'username': os.getenv('VMS_USERNAME_1'),
        'password': os.getenv('VMS_PASSWORD_1'),
        'org_key': os.getenv('VMS_ORG_KEY_1')
    },
    {
        'username': os.getenv('VMS_USERNAME_2'),
        'password': os.getenv('VMS_PASSWORD_2'), 
        'org_key': os.getenv('VMS_ORG_KEY_2')
    }
]

# Email configuration - USE ENVIRONMENT VARIABLES FOR SECURITY
EMAIL_CONFIG = {
    'sender_email': os.getenv('EMAIL_USER'),
    'password': os.getenv('EMAIL_PASSWORD'),
    'smtp_server': "smtp.gmail.com",
    'smtp_port': 587,
    'to_emails': ["jobdescriptions1@gmail.com"],   # Multiple TO recipients  cathie@iitlabs.com   jobdescriptions1@gmail.com
    'cc_emails': [],              # "support@innosoul.com"
    'bcc_emails': []
}

STATE_MAPPING = {
    'VA': {
        'full_name': 'Virginia',
        'city': 'Richmond',
        'sm_template': 'SM_Virginia.docx',
        'rtr_template': 'RTR_Virginia.docx'
    },
    'NC': {
        'full_name': 'North Carolina',
        'city': 'Raleigh', 
        'sm_template': 'SM_North_Carolina.docx',
        'rtr_template': 'RTR_North_Carolina.docx'
    },
    'GA': {
        'full_name': 'Georgia',
        'city': 'Atlanta',
        'sm_template': 'SM_Georgia.docx',
        'rtr_template': 'RTR_Georgia.docx'
    },
    'IN': {
        'full_name': 'Indiana',
        'city': 'Indianapolis',
        'sm_template': 'SM_Indiana.docx',
        'rtr_template': 'RTR_Indiana.docx'
    },
    'FL': {
        'full_name': 'Florida',
        'city': 'Jacksonville',
        'sm_template': 'SM_Florida.docx',
        'rtr_template': 'RTR_Florida.docx'
    },
    'ID': {
        'full_name': 'Idaho',
        'city': 'Boise',
        'sm_template': 'RTR_SM_Idaho.docx',
        'rtr_template': 'RTR_SM_Idaho.docx'
    },
    'IA': {
        'full_name': 'Iowa',
        'city': 'Cedar Rapids',
        'sm_template': 'SM_Iowa.docx',
        'rtr_template': 'RTR_Iowa.docx'
    },
    'DEFAULT': {
        'full_name': 'Default',
        'city': 'Default',
        'sm_template': 'SM.docx',
        'rtr_template': 'RTR.docx'
    }
}

def extract_state_from_job_id(content):
    """Extract state abbreviation from Job ID line in content with proper prioritization"""
    lines = content.split('\n')
    
    # PRIORITY 1: Look for Job ID line pattern in the FIRST FEW LINES (most reliable)
    for i, line in enumerate(lines[:10]):  # Only check first 10 lines
        if line.startswith('Job ID:') or 'Job ID:' in line:
            match = re.search(r'Job ID:\s*([A-Z]{2})-\d+', line)
            if match:
                state_abbr = match.group(1)
                if state_abbr in STATE_MAPPING:
                    print(f"  Extracted state from Job ID line: {state_abbr}")
                    return state_abbr
    
    # PRIORITY 2: Look for state pattern in the TITLE (first few lines)
    for i, line in enumerate(lines[:5]):  # Check first 5 lines for title
        # Look for patterns like: FL-DOT-, NC-FAST-, VA-123, etc.
        match = re.search(r'^([A-Z]{2})-[A-Z]', line)
        if match:
            state_abbr = match.group(1)
            if state_abbr in STATE_MAPPING:
                print(f"  Extracted state from title pattern: {state_abbr}")
                return state_abbr
    
    # PRIORITY 3: Look for state abbreviations in the title or content (avoid template text)
    for line in lines:
        # Skip lines that look like they're from templates, not actual job data
        if any(template_text in line for template_text in 
               ['VectorVMS Requirement', 'Candidate Full Legal Name', 'Candidate Pay Rate']):
            continue
            
        for state_abbr in STATE_MAPPING.keys():
            if state_abbr != 'DEFAULT' and state_abbr in line:
                # Make sure it's not part of a word and is a valid state code
                if re.search(r'\b' + state_abbr + r'\b', line):
                    print(f"  Extracted state from content: {state_abbr}")
                    return state_abbr
    
    # PRIORITY 4: Look for city names in content (avoid template cities)
    for line in lines:
        # Skip template-looking lines
        if 'Candidate' in line or 'VectorVMS' in line:
            continue
            
        for state_abbr, state_info in STATE_MAPPING.items():
            if state_abbr != 'DEFAULT' and state_info['city'].lower() in line.lower():
                print(f"  Extracted state from city: {state_abbr}")
                return state_abbr
    
    # PRIORITY 5: Look for full state names in content (avoid template text)
    for line in lines:
        if 'Managed Services Provider Contract' in line:
            continue  # Skip template text
            
        for state_abbr, state_info in STATE_MAPPING.items():
            if state_abbr != 'DEFAULT' and state_info['full_name'].lower() in line.lower():
                print(f"  Extracted state from full name: {state_abbr}")
                return state_abbr
    
    print("  Could not determine state, using DEFAULT")
    return 'DEFAULT'  # Final fallback

def validate_email_addresses(email_list):
    """Validate email addresses before sending"""
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    valid_emails = []
    
    for email in email_list if email_list else []:
        if isinstance(email, str) and re.match(email_regex, email):
            valid_emails.append(email)
        else:
            logging.warning(f"Invalid email address: {email}")
    
    return valid_emails

def attach_file(msg, file_path, filename):
    """Helper function to attach files with error handling"""
    if os.path.exists(file_path):
        try:
            with open(file_path, 'rb') as file:
                part = MIMEApplication(file.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)
            return True
        except Exception as e:
            logging.error(f"Failed to attach {filename}: {str(e)}")
            return False
    else:
        logging.warning(f"File not found for attachment: {file_path}")
        return False

def send_requisition_email(req_id, title, state_abbr, body_content, sm_path, rtr_path, 
                          to_emails=None, cc_emails=None, bcc_emails=None):
    """Send email with state-specific requisition content and documents"""
    try:
        # Validate email addresses
        to_emails = validate_email_addresses(to_emails or EMAIL_CONFIG.get('to_emails', []))
        cc_emails = validate_email_addresses(cc_emails or EMAIL_CONFIG.get('cc_emails', []))
        bcc_emails = validate_email_addresses(bcc_emails or EMAIL_CONFIG.get('bcc_emails', []))
        
        if not any([to_emails, cc_emails, bcc_emails]):
            logging.error("No valid email recipients found")
            return False
        
        sender_email = EMAIL_CONFIG['sender_email']
        password = EMAIL_CONFIG['password']
        smtp_server = EMAIL_CONFIG['smtp_server']
        smtp_port = EMAIL_CONFIG['smtp_port']
        
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = sender_email
        
        # Set recipients
        if to_emails:
            msg['To'] = ', '.join(to_emails)
        
        # Add CC if specified
        if cc_emails:
            msg['Cc'] = ', '.join(cc_emails)
        
        # BCC is not added to headers (handled separately in sendmail)
        state_name = STATE_MAPPING.get(state_abbr, STATE_MAPPING['DEFAULT'])['full_name']
        msg['Subject'] = f"{title}"
        
        # Email body with state info
        body = f""" {body_content}

Please find attached the SM and RTR documents for this {state_name} requisition.
"""
        msg.attach(MIMEText(body, 'plain'))
        
        # ENHANCED: Handle combined documents (like Idaho)
        sm_filename = os.path.basename(sm_path)
        rtr_filename = os.path.basename(rtr_path)
        
        # Check if this is a combined document (same file for both SM and RTR)
        if sm_path == rtr_path and os.path.exists(sm_path):
            print(f"  📧 Attaching combined document as both SM and RTR")
            # Attach the same file twice with different names
            attach_file(msg, sm_path, f"SM_{state_abbr}_{req_id}.docx")
            attach_file(msg, rtr_path, f"RTR_{state_abbr}_{req_id}.docx")
        else:
            # Separate documents - attach normally
            attach_file(msg, sm_path, f"SM_{state_abbr}_{req_id}.docx")
            attach_file(msg, rtr_path, f"RTR_{state_abbr}_{req_id}.docx")
        
        # Prepare all recipients (TO + CC + BCC)
        all_recipients = []
        if to_emails:
            all_recipients.extend(to_emails)
        if cc_emails:
            all_recipients.extend(cc_emails)
        if bcc_emails:
            all_recipients.extend(bcc_emails)
        
        # Remove duplicates
        all_recipients = list(set(all_recipients))
        
        # Send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, all_recipients, msg.as_string())
        
        # Log recipients
        recipient_info = f"TO: {to_emails}" if to_emails else ""
        if cc_emails:
            recipient_info += f", CC: {cc_emails}"
        if bcc_emails:
            recipient_info += f", BCC: {bcc_emails}"
        
        logging.info(f"Email sent successfully for {state_name} requisition {req_id} to {recipient_info}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send email for {state_abbr} requisition {req_id}: {str(e)}")
        return False

def auto_authenticate_primary_gmail():
    """Authenticates with primary Gmail account (token.json)"""
    print("\n=== Authenticating Primary Gmail Account ===")
    creds = None
    token_path = 'token.json'
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
   
    if os.path.exists(token_path):
        print("Found primary token file, loading credentials...")
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
            if creds.expired and creds.refresh_token:
                print("Primary credentials expired, refreshing...")
                creds.refresh(Request())
        except Exception as e:
            print(f"Error loading primary credentials: {e}")
            creds = None
   
    if not creds or not creds.valid:
        print("No valid primary credentials found, initiating OAuth flow...")
        try:
            flow = InstalledAppFlow.from_client_secrets_file('client.json', SCOPES)
            creds = flow.run_local_server(port=8080)
            token_data = json.loads(creds.to_json())
            token_data['creation_time'] = datetime.now(pytz.UTC).isoformat()
            with open(token_path, 'w') as token:
                json.dump(token_data, token)
            print("Primary authentication successful! Token saved.")
        except Exception as e:
            print(f"Primary authentication failed: {e}")
            raise
   
    try:
        print("Building primary Gmail service...")
        gmail_service = build('gmail', 'v1', credentials=creds)
        print("Primary Gmail service ready!")
        return gmail_service
    except Exception as e:
        print(f"Failed to build primary Gmail service: {e}")
        raise

def extract_direct_links(html_body, plain_text):
    """Extracts only Vector VMS links and ignores other links"""
    links = []
   
    # Pattern to match Vector VMS links
    vms_pattern = r'https?://vms\.vectorvms\.com[^\s<>"]+'
   
    # Extract from HTML
    if html_body:
        soup = BeautifulSoup(html_body, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if re.search(vms_pattern, href):
                links.append(href)
   
    # Extract from plain text
    if plain_text:
        text_links = re.findall(vms_pattern, plain_text)
        links.extend(text_links)
   
        # Also capture links after specific text patterns
        prompt_links = re.findall(
            r'(?:Click link to access requisition information|Click here|here):?\s*(https?://vms\.vectorvms\.com[^\s<>"]+)',
            plain_text,
            re.IGNORECASE
        )
        links.extend(prompt_links)
   
    return list(set(links))
 
def login_to_vectorvms(driver, url, credentials, max_retries=3):
    """Logs into Vector VMS system with enhanced reliability"""
    print(f"  Trying org_key: {credentials['org_key']}")
    driver.get(url)
    time.sleep(5)  # Initial wait for page load
 
    for attempt in range(max_retries):
        try:
            # Wait for login page to be fully interactive
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body')))
           
            # Explicitly wait for and locate login fields
            username_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                    "//input[@type='text' or contains(@id, 'username') or contains(@name, 'username') or contains(@placeholder, 'Username')]")))
            password_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                    "//input[@type='password' or contains(@id, 'password') or contains(@name, 'password')]")))
            org_key_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                    "//input[contains(@id, 'org') or contains(@name, 'org') or contains(@placeholder, 'Organization Key')]")))
 
            # Fill credentials - USING EXACT VALUES FROM CREDENTIAL_SETS
            username_field.clear()
            username_field.send_keys(credentials['username'])  # 'support'
            password_field.clear()
            password_field.send_keys(credentials['password'])  # 'db3admin'
            org_key_field.clear()
            org_key_field.send_keys(credentials['org_key'])    # org_key from list
 
            # Click login button with improved specificity
            login_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in')] | " +
                    "//input[@type='submit' and contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login')]")))
            login_button.click()
           
            # Wait for successful login or requisition page
            WebDriverWait(driver, 40).until(
                lambda d: "reqid" in d.current_url.lower() or "dashboard" in d.current_url.lower(),
                message="Failed to reach requisition or dashboard page")
           
            print(f"  ✓ Login successful with org_key: {credentials['org_key']}!")
            return True
 
        except Exception as e:
            print(f"  Login attempt {attempt + 1}/{max_retries} failed: {str(e)}")
            if attempt < max_retries - 1:
                print("  Retrying in 5 seconds...")
                time.sleep(5)
                driver.get(url)  # Reload page for retry
            else:
                print(f"  ✗ All login attempts failed with org_key: {credentials['org_key']}")
                return False

def logout_from_vectorvms(driver):
    """Logs out from Vector VMS system"""
    print("  Logging out from Vector VMS...")
    try:
        # Try to find and click logout button
        logout_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, 
                "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'logout') or " +
                "contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign out')]")))
        logout_button.click()
        
        # Wait for logout to complete (redirect to login page)
        WebDriverWait(driver, 15).until(
            lambda d: "login" in d.current_url.lower() or "signin" in d.current_url.lower())
        
        print("  ✓ Logout successful!")
        return True
    except Exception as e:
        print(f"  Logout failed: {str(e)}")
        # If logout fails, try to clear cookies and refresh
        try:
            driver.delete_all_cookies()
            driver.refresh()
            print("  Cleared cookies as fallback logout method")
            return True
        except:
            print("  Could not perform clean logout")
        return False
 
def extract_complete_page_content(driver):
    """Extracts all visible content from the page with dynamic key-value pairs, including both short and complete descriptions"""
 
    try:
        # Wait for the main content to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body')))
 
        # Scroll through the entire page to load all content
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
 
        # Switch to main iframe if present (save the element for later if needed)
        main_iframe = None
        try:
            main_iframe = driver.find_element(By.TAG_NAME, "iframe")
            driver.switch_to.frame(main_iframe)
            print("  Switched to main iframe context")
        except:
            print("  No main iframe detected")
 
        # Get the page source and parse with BeautifulSoup
        page_html = driver.page_source
        soup = BeautifulSoup(page_html, 'html.parser')
        content = []
        seen_keys = set()
 
        # Handle section headers
        for header in soup.find_all('h2'):
            header_text = header.find('span', class_='x-panel-header-text')
            if header_text and header_text.text.strip() not in seen_keys:
                content.append(header_text.text.strip())
                seen_keys.add(header_text.text.strip())
 
        # Extract key-value pairs dynamically
        for label in soup.find_all(['label', 'span'], string=True):
            key = label.get_text(strip=True).rstrip(':')
            if key and key not in seen_keys and not any(kw in key for kw in ['label', 'header']):
                seen_keys.add(key)
                # Find the next value element
                next_elem = label.find_next(['input', 'textarea', 'span', 'div'],
                                          class_=lambda c: c and ('x-label-value' in c or 'vms-viewmode-view-set' in c or 'x-form-display-field' in c))
                if next_elem:
                    try:
                        element = driver.find_element(By.ID, next_elem.get('id'))
                        if element.tag_name == 'span':
                            value = element.text.strip()
                        elif element.tag_name in ['input', 'textarea']:
                            value = driver.execute_script("return arguments[0].value", element) or element.text.strip()
                        elif element.tag_name == 'div' and 'x-form-display-field' in element.get_attribute('class'):
                            value = element.text.strip()
                        else:
                            value = next_elem.get_text(strip=True)
                        if value and len(value) > 0:
                            content.append(f"{key}: {value}")
                        else:
                            content.append(f"{key}:")
                    except:
                        content.append(f"{key}:")
                else:
                    content.append(f"{key}:")
 
        # Extract Short Description
        short_desc = ""
        try:
            # Try to find the short description section
            short_desc_div = soup.find('div', id='ContentPH_ctl209')
            if short_desc_div:
                # Try to get content from iframe
                iframe = short_desc_div.find('iframe')
                if iframe:
                    driver.switch_to.frame(iframe.get('name'))
                    iframe_html = driver.page_source
                    iframe_soup = BeautifulSoup(iframe_html, 'html.parser')
                    short_desc = iframe_soup.body.get_text('\n', strip=True)
                    driver.switch_to.parent_frame()  # Switch back to parent context (main iframe)
                else:
                    # Fallback to hidden textarea
                    hidden_textarea = short_desc_div.find('textarea', id='ContentPH_description_short_html')
                    if hidden_textarea and hidden_textarea.text.strip():
                        short_desc = hidden_textarea.text.strip()
                    else:
                        # Final fallback to display field
                        display_field = short_desc_div.find('div', id='ContentPH_lblShortDescription')
                        if display_field and display_field.text.strip():
                            short_desc = display_field.text.strip()
        except Exception as e:
            print(f"  Short description extraction error: {str(e)}")
            try:
                driver.switch_to.parent_frame()  # Use parent_frame in except too
            except:
                pass
 
        # Add short description if found
        if short_desc and short_desc.strip():
            content.append("\nSHORT DESCRIPTION:")
            content.append(short_desc.strip())
 
        # Extract Complete Description
        complete_desc = ""
        try:
            # Try to find the complete description section
            complete_desc_div = soup.find('div', id='ContentPH_ctl214')
            if complete_desc_div:
                # Try to get content from iframe
                iframe = complete_desc_div.find('iframe')
                if iframe:
                    driver.switch_to.frame(iframe.get('name'))
                    iframe_html = driver.page_source
                    iframe_soup = BeautifulSoup(iframe_html, 'html.parser')
                    complete_desc = iframe_soup.body.get_text('\n', strip=True)
                    driver.switch_to.parent_frame()  # Switch back to parent context (main iframe)
                else:
                    # Fallback to hidden textarea
                    hidden_textarea = complete_desc_div.find('textarea', id='ContentPH_description_html')
                    if hidden_textarea and hidden_textarea.text.strip():
                        complete_desc = hidden_textarea.text.strip()
                    else:
                        # Final fallback to display field
                        display_field = complete_desc_div.find('div', id='ContentPH_lblDescription')
                        if display_field and display_field.text.strip():
                            complete_desc = display_field.text.strip()
            # Additional fallback: search for complete description in other divs or text
            if not complete_desc:
                for div in soup.find_all('div', class_=['job-description', 'description', 'complete-description']):
                    text = div.get_text('\n', strip=True)
                    if text and 'description' in div.get('id', '').lower():
                        complete_desc = text
                        break
                if not complete_desc:
                    for p in soup.find_all('p', string=True):
                        text = p.get_text(strip=True)
                        if text and any(kw in text.lower() for kw in ['job description', 'complete description', 'requisition description']):
                            complete_desc = text
                            break
        except Exception as e:
            print(f"  Complete description extraction error: {str(e)}")
            try:
                driver.switch_to.parent_frame()  # Use parent_frame in except too
            except:
                pass
 
        # Add complete description if found
        if complete_desc and complete_desc.strip():
            content.append("\nCOMPLETE DESCRIPTION:")
            content.append(complete_desc.strip())
 
        # Capture Requisition Description and other standalone text
        description_elements = soup.find_all(['p', 'body'], string=True)
        for elem in description_elements:
            text = elem.get_text(strip=True)
            if text and len(text) > 2 and not text.isspace() and text not in seen_keys:
                if any(kw in text for kw in ['JOB DESCRIPTION', 'KNOWLEDGE, SKILLS, AND ABILITIES']):
                    content.append(text)
                elif 'Contract' in text:
                    content.append("Engagement Type: Contract")
                else:
                    content.append(text)
 
        # Capture Work Location from div with specific class
        work_location = soup.find('div', class_='ux-mselect-item')
        if work_location and work_location.text.strip() not in seen_keys:
            content.append(f"Work Location: {work_location.text.strip()}")
 
        # Switch back to main content if needed
        driver.switch_to.default_content()  # Final reset to top level
 
        return '\n'.join(content)
 
    except Exception as e:
        print(f"  Extraction error: {str(e)}")
        driver.switch_to.default_content()  # Final reset in error
        return driver.page_source

def initialize_driver():
    """Initialize and return a Chrome WebDriver with optimal settings"""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def wait_for_stable_element(driver, locator, timeout=30, stability_time=2):
    """Wait for element to be present and stable (not changing)"""
    end_time = time.time() + timeout
    last_html = ""
    
    while time.time() < end_time:
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(locator)
            )
            current_html = element.get_attribute('outerHTML')
            if current_html == last_html:
                time.sleep(stability_time)  # Additional stability wait
                return element
            last_html = current_html
            time.sleep(0.5)
        except:
            time.sleep(0.5)
    
    raise TimeoutError(f"Element not stable after {timeout} seconds")

def cleanup_memory(driver):
    """Execute JavaScript to clean up memory"""
    try:
        driver.execute_script("window.gc();")  # Trigger garbage collection
    except:
        pass

def extract_skills_table(driver, max_retries=3):
    """Robust skills table extraction with comprehensive error handling and restart after access"""
    
    def safe_get_text(element, default="N/A"):
        """Helper function to safely get text from an element"""
        try:
            text = element.text.strip()
            return text if text else default
        except:
            return default

    for attempt in range(max_retries):
        try:
            print(f"  Attempt {attempt + 1} of {max_retries} to extract skills table")
            cleanup_memory(driver)
            
            # Click Skills tab with multiple strategies
            try:
                skills_button = wait_for_stable_element(
                    driver, 
                    (By.XPATH, "//a[contains(., 'Skills')]"),
                    timeout=20
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", skills_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", skills_button)
                
                # ADDED: Restart/refresh logic specifically for skills section
                print("  Restarting skills section after access...")
                time.sleep(5)  # Wait 5 seconds before refresh
                driver.refresh()  # Refresh the page
                
                # Re-locate and click skills tab after refresh
                skills_button = wait_for_stable_element(
                    driver, 
                    (By.XPATH, "//a[contains(., 'Skills')]"),
                    timeout=20
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", skills_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", skills_button)
                
            except Exception as click_error:
                print(f"  Skills tab click failed: {click_error}. Trying URL navigation fallback...")
                current_url = driver.current_url
                if "reqID=" in current_url and "#" not in current_url:
                    driver.get(current_url + "#skills")
                    # ADDED: Restart/refresh logic for URL navigation approach
                    print("  Restarting skills section after URL navigation...")
                    time.sleep(5)
                    driver.refresh()
                    driver.get(current_url + "#skills")
                time.sleep(3)

            # Handle iframes if present
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    print("  Found iframe, switching context...")
                    driver.switch_to.frame(iframes[0])
            except:
                print("  No iframe detected or error switching to iframe")

            # Wait for skills grid to load completely
            try:
                wait_for_stable_element(
                    driver,
                    (By.CSS_SELECTOR, "div.x-grid3-body"),  # FIXED: CSS_SELECTOR not CSS_SELECTor
                    timeout=25
                )
                time.sleep(1)  # Additional stabilization wait
            except Exception as wait_error:
                print(f"  Waiting for skills grid failed: {wait_error}")
                raise

            # Extract table data
            rows = []
            row_elements = driver.find_elements(By.CSS_SELECTOR, "div.x-grid3-row")  # FIXED: CSS_SELECTOR
            
            if not row_elements:
                print("  No skill rows found, checking for alternative structure...")
                row_elements = driver.find_elements(By.CSS_SELECTOR, "table.x-grid3-row-table")  # FIXED: CSS_SELECTOR
                
            for i, row in enumerate(row_elements):
                try:
                    # Get all cells in the row
                    cells = row.find_elements(By.CSS_SELECTOR, "td.x-grid3-col")  # FIXED: CSS_SELECTOR
                    
                    # Extract skill (first column)
                    skill = safe_get_text(cells[1].find_element(By.CSS_SELECTOR, "span.x-grid3-cell-inner"))  # FIXED: CSS_SELECTOR
                    
                    # Extract type (second column)
                    req_desired = safe_get_text(cells[2].find_element(By.CSS_SELECTOR, "span.x-grid3-cell-inner"))  # FIXED: CSS_SELECTOR
                    
                    # Extract duration (third column)
                    amount = safe_get_text(cells[3].find_element(By.CSS_SELECTOR, "span.x-grid3-cell-inner"))  # FIXED: CSS_SELECTOR
                    
                    # Extract duration type (fourth column)
                    duration_type = safe_get_text(cells[4].find_element(By.CSS_SELECTOR, "span.x-grid3-cell-inner"))  # FIXED: CSS_SELECTOR
                    
                    # Format experience
                    if amount == "N/A" or duration_type == "N/A":
                        experience = "N/A"
                    else:
                        experience = f"{amount} {duration_type}".strip()
                    
                    rows.append([skill, req_desired, experience])
                    
                except Exception as row_error:
                    print(f"  Error processing row {i + 1}: {row_error}")
                    continue

            # Format results as markdown table
            headers = ["Skill", "Type", "Experience"]
            markdown_table = [
                "| " + " | ".join(headers) + " |",
                "| " + " | ".join(["---"] * len(headers)) + " |"
            ]
            
            for row in rows:
                clean_row = [cell.replace("\n", " ").strip() for cell in row]
                markdown_table.append("| " + " | ".join(clean_row) + " |")
            
            print("  Skills table extracted successfully")
            return "\n".join(markdown_table)

        except Exception as e:
            print(f"  Attempt {attempt + 1} failed with error: {str(e)}")
            
            # Capture debugging info
            try:
                print(f"  Current URL: {driver.current_url}")
                print(f"  Page title: {driver.title}")
                driver.save_screenshot(f"error_attempt_{attempt + 1}.png")
            except:
                pass
            
            if attempt < max_retries - 1:
                # Recovery actions
                try:
                    driver.switch_to.default_content()
                    driver.refresh()
                    time.sleep(5)
                except:
                    # If recovery fails, restart browser
                    try:
                        driver.quit()
                    except:
                        pass
                    driver = initialize_driver()
            else:
                print("  Max retries reached, returning empty table")
                return "| Skill | Type | Experience |\n|-------|------|------------|\n| N/A | N/A | N/A |"
    
    return "| Skill | Type | Experience |\n|-------|------|------------|\n| N/A | N/A | N/A |"

def extract_requisition_urls_from_gmail(gmail_service):
    """Extracts Vector VMS requisition URLs from TODAY'S Gmail emails with 'Now Open' subject"""
    print("\n=== Extracting TODAY'S Requisition URLs from Gmail ===")
    
    try:
        # Get today's date in the format Gmail API expects (YYYY/MM/DD)
        today_date = datetime.now().strftime("%Y/%m/%d")
        print(f"Looking for emails from today: {today_date}")
        
        # Search for TODAY'S emails with "Now Open" in subject
        query = f'subject:"Now Open" newer_than:1d'
        print(f"Gmail query: {query}")
        
        results = gmail_service.users().messages().list(
            userId='me',
            labelIds=['INBOX'],
            q=query
        ).execute()
        messages = results.get('messages', [])
        
        if not messages:
            print("No 'Now Open' emails found from the last day.")
            # CHANGED: Don't fall back to searching without date filter
            print("No TODAY'S requisition emails found. Stopping processing.")
            return []  # Return empty list instead of falling back
        
        print(f"Found {len(messages)} 'Now Open' emails from today. Scanning for Vector VMS links...")
        
        # Collect TODAY'S links only
        today_links = set()
        processed_count = 0
        today_count = 0
        
        for i, msg in enumerate(messages, 1):
            print(f"Scanning email {i}/{len(messages)}")
            try:
                msg_data = gmail_service.users().messages().get(
                    userId='me',
                    id=msg['id'],
                    format='full'
                ).execute()
                
                # Get the internal timestamp (more reliable than header date)
                internal_date = int(msg_data['internalDate'])
                email_date = datetime.fromtimestamp(internal_date / 1000)
                is_today = email_date.date() == datetime.now().date()
                
                if is_today:
                    today_count += 1
                    print(f"  ✓ Email from today: {email_date}")
                else:
                    print(f"  ⚠ Email from {email_date.date()} (not today) - SKIPPING")
                    continue  # CHANGED: Skip emails that are not from today
                
                payload = msg_data['payload']
                html_body = ""
                plain_text = ""
                
                # Extract email content
                if 'parts' in payload:
                    for part in payload['parts']:
                        if part['mimeType'] == 'text/html':
                            data = part['body'].get('data', '')
                            if data:
                                html_body = base64.urlsafe_b64decode(data).decode('utf-8')
                        elif part['mimeType'] == 'text/plain':
                            data = part['body'].get('data', '')
                            if data:
                                plain_text = base64.urlsafe_b64decode(data).decode('utf-8')
                else:
                    if payload['mimeType'] == 'text/html':
                        data = payload['body'].get('data', '')
                        if data:
                            html_body = base64.urlsafe_b64decode(data).decode('utf-8')
                    elif payload['mimeType'] == 'text/plain':
                        data = payload['body'].get('data', '')
                        if data:
                            plain_text = base64.urlsafe_b64decode(data).decode('utf-8')

                links = extract_direct_links(html_body, plain_text)
                today_links.update(links)
                
                processed_count += 1
                
            except Exception as e:
                print(f"Error scanning email: {e}")
                continue
        
        print(f"\nProcessed {processed_count} emails from today")
        print(f"Found {len(today_links)} Vector VMS links from today's emails")
        
        # CHANGED: Only return today's links, never fall back to old links
        if today_links:
            return list(today_links)
        else:
            print("No Vector VMS links found in today's emails.")
            return []
        
    except Exception as e:
        print(f"Error extracting URLs from Gmail: {e}")
        return []

# ===== DOCUMENT PROCESSING FUNCTIONS =====


def extract_state_from_job_id(content):
    """Extract state abbreviation from Job ID line in content with proper prioritization"""
    lines = content.split('\n')
    
    # PRIORITY 1: Look for Job ID line pattern in the FIRST FEW LINES (most reliable)
    for i, line in enumerate(lines[:10]):  # Only check first 10 lines
        if line.startswith('Job ID:') or 'Job ID:' in line:
            match = re.search(r'Job ID:\s*([A-Z]{2})-\d+', line)
            if match:
                state_abbr = match.group(1)
                if state_abbr in STATE_MAPPING:
                    print(f"  Extracted state from Job ID line: {state_abbr}")
                    return state_abbr
    
    # PRIORITY 2: Look for state pattern in the TITLE (first few lines)
    for i, line in enumerate(lines[:5]):  # Check first 5 lines for title
        # Look for patterns like: FL-DOT-, NC-FAST-, VA-123, etc.
        match = re.search(r'^([A-Z]{2})-[A-Z]', line)
        if match:
            state_abbr = match.group(1)
            if state_abbr in STATE_MAPPING:
                print(f"  Extracted state from title pattern: {state_abbr}")
                return state_abbr
    
    # PRIORITY 3: Look for state abbreviations in the title or content (avoid template text)
    for line in lines:
        # Skip lines that look like they're from templates, not actual job data
        if any(template_text in line for template_text in 
               ['VectorVMS Requirement', 'Candidate Full Legal Name', 'Candidate Pay Rate']):
            continue
            
        for state_abbr in STATE_MAPPING.keys():
            if state_abbr != 'DEFAULT' and state_abbr in line:
                # Make sure it's not part of a word and is a valid state code
                if re.search(r'\b' + state_abbr + r'\b', line):
                    print(f"  Extracted state from content: {state_abbr}")
                    return state_abbr
    
    # PRIORITY 4: Look for city names in content (avoid template cities)
    for line in lines:
        # Skip template-looking lines
        if 'Candidate' in line or 'VectorVMS' in line:
            continue
            
        for state_abbr, state_info in STATE_MAPPING.items():
            if state_abbr != 'DEFAULT' and state_info['city'].lower() in line.lower():
                print(f"  Extracted state from city: {state_abbr}")
                return state_abbr
    
    # PRIORITY 5: Look for full state names in content (avoid template text)
    for line in lines:
        if 'Managed Services Provider Contract' in line:
            continue  # Skip template text
            
        for state_abbr, state_info in STATE_MAPPING.items():
            if state_abbr != 'DEFAULT' and state_info['full_name'].lower() in line.lower():
                print(f"  Extracted state from full name: {state_abbr}")
                return state_abbr
    
    print("  Could not determine state, using DEFAULT")
    return 'DEFAULT'  # Final fallback

def parse_skills_table(skills_content):
    """
    Parse the skills content - handles both markdown table and bullet formats
    """
    skills_data = []
    
    if not skills_content:
        print("  No skills content provided to parse")
        return skills_data
    
    print(f"  Parsing skills content: {len(skills_content)} chars")
    
    # Check if it's a placeholder table
    if "N/A | N/A | N/A" in skills_content:
        print("  Placeholder skills table found")
        return skills_data
    
    # Check if it's regex-extracted format (SKILLS TABLE: with bullet points)
    if "SKILLS TABLE:" in skills_content or (('Required' in skills_content or 'Desired' in skills_content) and 'Years' in skills_content and '|' not in skills_content):
        print("  🔍 Detected regex-extracted skills format")
        return parse_regex_extracted_skills(skills_content)
    
    # Otherwise, try to parse as markdown table
    lines = skills_content.split('\n')
    print(f"  Total lines in skills table: {len(lines)}")
    
    # Find the data rows (skip header and separator)
    data_lines = []
    in_data_section = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Markdown table separator (|---|)
        if line.startswith('|---') or '--- | ---' in line or line.startswith('|--'):
            in_data_section = True
            continue
        
        # Data rows (start with | and not header)
        if line.startswith('|') and in_data_section:
            if 'Skill' not in line and '---' not in line:
                data_lines.append(line)
    
    print(f"  Found {len(data_lines)} data rows to process")
    
    # Process each data row
    for line in data_lines:
        try:
            # Split by pipe and clean up
            parts = [part.strip() for part in line.split('|')]
            # Remove empty first and last parts (from leading/trailing |)
            parts = [p for p in parts if p]
            
            if len(parts) >= 3:
                skill = parts[0]
                skill_type = parts[1]
                experience = parts[2]
                
                # Extract years from experience with better parsing
                years = "0"
                if experience and experience != "N/A":
                    # Look for numbers in experience
                    years_match = re.search(r'(\d+\.?\d*)\s*(?:year|yr|years|y)\w*', experience, re.IGNORECASE)
                    if years_match:
                        years = years_match.group(1)
                    else:
                        # Try to find any number
                        any_num_match = re.search(r'(\d+)', experience)
                        if any_num_match:
                            years = any_num_match.group(1)
                
                # Map to proper type for VA template
                requirement_type = "Required"
                if 'desired' in skill_type.lower() or 'preferred' in skill_type.lower():
                    requirement_type = "Desired"
                elif 'highly' in skill_type.lower():
                    requirement_type = "Highly desired"
                
                skills_data.append({
                    'skill': skill,
                    'type': requirement_type,
                    'experience': experience,
                    'years': years
                })
                print(f"    ✅ Found skill: {skill} - {requirement_type} - {experience}")
                
            elif len(parts) == 1:
                # Just a skill name without type/experience
                skill = parts[0]
                skills_data.append({
                    'skill': skill,
                    'type': 'Required',
                    'experience': '',
                    'years': '0'
                })
                print(f"    ✅ Found skill: {skill} (no type/experience)")
                
        except Exception as e:
            print(f"    ❌ Error processing skill row: {e}")
            continue
    
    print(f"  Total skills parsed: {len(skills_data)}")
    
    return skills_data


def parse_regex_extracted_skills(skills_text):
    """
    Parse skills from regex-extracted format like:
    SKILLS TABLE:
    Manage vendor relationships... Required 3 Years
    Experience tracking SLAs... Required 3 Years
    """
    skills_data = []
    
    if not skills_text:
        return skills_data
    
    print(f"  📝 Parsing regex-extracted skills format")
    
    # Remove "SKILLS TABLE:" header if present
    if skills_text.startswith("SKILLS TABLE:"):
        skills_text = skills_text.replace("SKILLS TABLE:", "").strip()
    
    lines = skills_text.strip().split('\n')
    
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue
        
        # Skip empty lines or section headers
        if line.lower() in ['skills:', 'skills table:']:
            continue
        
        # Initialize defaults
        skill_desc = line
        req_type = "Required"
        years = "0"
        
        try:
            # Try to extract the pattern: "... Required 3 Years"
            # Look for "Years" at the end of the line
            if " Years" in line:
                # Find the last occurrence of "Years"
                years_idx = line.rfind(" Years")
                if years_idx != -1:
                    # Get the part before "Years"
                    before_years = line[:years_idx].strip()
                    # Get the part after "Years" (should be empty or continuation)
                    after_years = line[years_idx + 6:].strip()
                    
                    # Extract the number before "Years"
                    # Go backwards from years_idx to find the number
                    # Look for pattern: "Required 3" or "Desired 5" etc.
                    type_year_match = re.search(r'(\bRequired\b|\bDesired\b|\bHighly desired\b|\bHIGHLY DESIRED\b)\s+(\d+)', before_years, re.IGNORECASE)
                    
                    if type_year_match:
                        req_type = type_year_match.group(1).strip()
                        years = type_year_match.group(2).strip()
                        
                        # Get skill description (everything before the requirement type)
                        req_type_start = before_years.rfind(req_type)
                        if req_type_start != -1:
                            skill_desc = before_years[:req_type_start].strip()
                        else:
                            skill_desc = before_years
                    
                    # Standardize requirement type
                    req_type_upper = req_type.upper()
                    if 'HIGHLY' in req_type_upper:
                        req_type = "HIGHLY DESIRED"
                    elif 'REQUIRED' in req_type_upper:
                        req_type = "Required"
                    elif 'DESIRED' in req_type_upper:
                        req_type = "Desired"
                    else:
                        req_type = "Required"
                        
                    # Clean up skill description
                    # Remove trailing dots
                    while skill_desc.endswith('..') or skill_desc.endswith('.'):
                        skill_desc = skill_desc.rstrip('.')
                    skill_desc = skill_desc.strip()
                    
            # Alternative pattern matching for cases where above doesn't work
            else:
                # Try regex pattern: "skill.. Required 3 Years"
                pattern = r'(.+?)(?:\.\.|\.)\s+(Required|Desired|Highly desired|HIGHLY DESIRED)\s+(\d+)\s+Years'
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    skill_desc = match.group(1).strip()
                    req_type = match.group(2).strip()
                    years = match.group(3).strip()
                    
                    # Standardize
                    if req_type.upper() == 'HIGHLY desired'.upper():
                        req_type = "HIGHLY DESIRED"
                    elif 'required' in req_type.lower():
                        req_type = "Required"
                    elif 'desired' in req_type.lower():
                        req_type = "Desired"
        
        except Exception as e:
            print(f"    ⚠️ Error parsing line {line_num}: {e}")
            # Keep defaults
        
        # Final cleanup
        skill_desc = skill_desc.rstrip('.').strip()
        
        skills_data.append({
            'skill': skill_desc,
            'type': req_type,
            'experience': f"{years} Years",
            'years': years
        })
    
    print(f"  ✅ Parsed {len(skills_data)} skills from regex format")
    
    # Debug: Show parsed skills
    if skills_data:
        print(f"  📋 First 3 parsed skills:")
        for i, skill in enumerate(skills_data[:3]):
            print(f"    {i+1}. '{skill['skill']}' - {skill['type']} - {skill['years']} years")
    
    return skills_data

def set_table_borders(table):
    """
    Set borders for all cells in a table to maintain the table format
    """
    try:
        tbl = table._tbl
        tblPr = tbl.tblPr
        
        # Add table borders
        tblBorders = parse_xml(r'''
            <w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>''')
        
        if tblPr is None:
            from docx.oxml.xmlchemy import OxmlElement
            tblPr = OxmlElement('w:tblPr')
            tbl.append(tblPr)
        
        existing_borders = tblPr.find(qn('w:tblBorders'))
        if existing_borders is not None:
            tblPr.remove(existing_borders)
        tblPr.append(tblBorders)
    except:
        # If border setting fails, just continue without borders
        pass

def debug_template_usage(documents_dir, state_abbr):
    """Debug function to see what templates are being used"""
    state_info = STATE_MAPPING.get(state_abbr, STATE_MAPPING['DEFAULT'])
    
    sm_template_path = os.path.join(documents_dir, state_info['sm_template'])
    rtr_template_path = os.path.join(documents_dir, state_info['rtr_template'])
    
    sm_default_path = os.path.join(documents_dir, STATE_MAPPING['DEFAULT']['sm_template'])
    rtr_default_path = os.path.join(documents_dir, STATE_MAPPING['DEFAULT']['rtr_template'])
    
    print(f"\n=== DEBUG TEMPLATE USAGE FOR {state_abbr} ===")
    print(f"Looking for SM template: {state_info['sm_template']} - Exists: {os.path.exists(sm_template_path)}")
    print(f"Looking for RTR template: {state_info['rtr_template']} - Exists: {os.path.exists(rtr_template_path)}")
    print(f"Default SM template: {STATE_MAPPING['DEFAULT']['sm_template']} - Exists: {os.path.exists(sm_default_path)}")
    print(f"Default RTR template: {STATE_MAPPING['DEFAULT']['rtr_template']} - Exists: {os.path.exists(rtr_default_path)}")
    
    # List all files in Documents directory
    print("\nAll files in Documents folder:")
    for file in os.listdir(documents_dir):
        if file.endswith('.docx'):
            print(f"  - {file}")

def update_rtr_document(req_id, title, state_abbr, documents_dir, output_path, requisition_content=None, skills_data=None):
    """
    Update the RTR document for the specific state with the requisition title
    ENHANCED: Handles combined documents like RTR_SM_Idaho.docx
    """
    # If state_abbr is DEFAULT or unclear, try to extract from requisition content
    if state_abbr == 'DEFAULT' and requisition_content:
        state_abbr = extract_state_from_job_id(requisition_content)
        print(f"  Extracted state from content: {state_abbr}")

    # Get state info - this defines what template we WANT to find
    state_info = STATE_MAPPING.get(state_abbr, STATE_MAPPING['DEFAULT'])
    desired_template_name = state_info['rtr_template']
    print(f"  Target state: {state_abbr}, Searching for template: {desired_template_name}")

    # Build the full desired path
    template_path = os.path.join(documents_dir, desired_template_name)
    
    # Check if the desired state template exists
    if os.path.exists(template_path):
        print(f"✅ Found exact state template: {desired_template_name}")
    else:
        print(f"❌ Desired template not found: {desired_template_name}")
        # Look for ANY template that matches the state abbreviation pattern
        all_files = os.listdir(documents_dir)
        
        # Filter out files that contain numbers (these are generated output files)
        template_files = [f for f in all_files if f.endswith('.docx') and not any(char.isdigit() for char in f)]
        print(f"  Available template files (no numbers): {[f for f in template_files if 'RTR' in f or 'SM' in f]}")
        
        # Look for template files that contain the state abbreviation
        matching_templates = []
        for file in template_files:
            file_upper = file.upper()
            # Match patterns like: RTR_FL.docx, SM_FL.docx, RTR_SM_Idaho.docx (NO NUMBERS)
            if (state_abbr.upper() in file_upper and
                not any(char.isdigit() for char in file)):
                matching_templates.append(file)
        
        if matching_templates:
            # Found a template for the correct state! Use the first match
            template_path = os.path.join(documents_dir, matching_templates[0])
            print(f"✅ Using found state-specific template: {matching_templates[0]}")
        else:
            # LAST RESORT: Use the default template
            default_template = STATE_MAPPING['DEFAULT']['rtr_template']
            template_path = os.path.join(documents_dir, default_template)
            if os.path.exists(template_path):
                print(f"⚠️  No state template found. Using default: {default_template}")
            else:
                print(f"❌ ERROR: No default template found either!")
                return None

    # Load the document
    try:
        doc = Document(template_path)
        print(f"✅ Loaded template: {os.path.basename(template_path)}")
    except Exception as e:
        print(f"❌ Error loading template {template_path}: {e}")
        return None

    # === FIX: UNIVERSAL SKILLS DATA HANDLING FOR ALL STATES ===
    # If skills_data is None but we have requisition_content, try to extract skills
    if skills_data is None and requisition_content:
        print("  ⚠️ No skills data provided, attempting to extract from content...")
        if "=== SKILLS TABLE ===" in requisition_content:
            try:
                skills_part = requisition_content.split("=== SKILLS TABLE ===")[1]
                skills_table_content = skills_part.split("\n\n")[0].strip()
                skills_data = parse_skills_table(skills_table_content)
                print(f"  ✅ Extracted {len(skills_data)} skills from requisition content")
            except Exception as e:
                print(f"  ❌ Failed to extract skills from content: {e}")
                skills_data = []
        else:
            print("  ❌ No skills table found in requisition content")
            skills_data = []
    elif skills_data is None:
        print("  ❌ No skills data available and no content to extract from")
        skills_data = []

    # Format the title consistently for both places
    formatted_title = f"{title} ({req_id})"
    state_name = state_info['full_name']
    print(f"  Updating RTR document for {state_name} with title: {formatted_title}")

    # Update all state references in the document - CAREFUL REPLACEMENT
    state_updated = False
    for paragraph in doc.paragraphs:
        original_text = paragraph.text
        
        # Only replace state references that are clearly template text, not job data
        is_template_text = any(pattern in original_text for pattern in [
            "Managed Services Provider Contract",
            "VectorVMS Requirement Number and Title",
            "Candidate Full Legal Name",
            "Candidate Pay Rate for this Position",
            "Candidate Employment Type if Selected for Engagement",
            "has the sole right to represent me in matters of work assignment"
        ])
        
        if is_template_text:
            # Replace state-specific content - Only replace if it's NOT the correct state
            for target_state_abbr, target_state_info in STATE_MAPPING.items():
                if target_state_abbr == 'DEFAULT' or target_state_abbr == state_abbr:
                    continue  # Skip DEFAULT and current state
                    
                # Replace full state names from other states (only in template text)
                if target_state_info['full_name'] in original_text:
                    new_text = original_text.replace(target_state_info['full_name'], state_name)
                    paragraph.text = new_text
                    state_updated = True
                    print(f"   Replaced '{target_state_info['full_name']}' with '{state_name}' in template text")
                
                # Replace state abbreviations from other states (only in template text)
                elif f" {target_state_abbr} " in f" {original_text} ":
                    new_text = original_text.replace(f" {target_state_abbr} ", f" {state_abbr} ")
                    paragraph.text = new_text
                    state_updated = True
                    print(f"   Replaced '{target_state_abbr}' with '{state_abbr}' in template text")
    
    if not state_updated:
        print("   ⚠️ No state-specific template content found to replace")
    
    # === FIX: TITLE CONSISTENCY ===
    # Update the title in email subject
    subject_updated = False
    email_subject_found = False
    for paragraph in doc.paragraphs:
        if "INSERT THE FOLLOWING INTO EMAIL SUBJECT" in paragraph.text:
            email_subject_found = True
        elif email_subject_found and paragraph.text.strip():
            # Found the subject line, update it with consistent title
            paragraph.text = formatted_title
            if paragraph.runs:
                paragraph.runs[0].font.size = Pt(14)
                paragraph.runs[0].bold = True
            print(f"   Updated email subject: {formatted_title}")
            subject_updated = True
            break
    
    # Update the title in Vector VMS section with the SAME title
    vms_updated = False
    vms_section_found = False
    for paragraph in doc.paragraphs:
        # KEEPING BOTH TEXT PATTERNS: "VectorVMS" AND "Vector VMS"
        if "VectorVMS Requirement Number and Title" in paragraph.text or "Vector VMS Requirement Number and Title" in paragraph.text:
            vms_section_found = True
        elif vms_section_found and paragraph.text.strip():
            # Found the VMS title, update it with the SAME consistent title
            paragraph.text = formatted_title
            if paragraph.runs:
                paragraph.runs[0].font.size = Pt(12)
            print(f"   Updated VMS title: {formatted_title}")
            vms_updated = True
            break

    # === FIX: UNIVERSAL SKILLS TABLE UPDATE FOR ALL STATES ===
    skills_updated_in_rtr = False
    
    # Check if we have skills data for ANY state (not just Florida)
    if skills_data and doc.tables:
        print(f"  Processing skills table for {state_abbr} document...")
        
        for table_idx, table in enumerate(doc.tables):
            # Look for ANY skills table with specific columns
            if len(table.rows) > 0:
                header_cells = table.rows[0].cells
                header_text = ' '.join(cell.text.strip().lower() for cell in header_cells)
                
                # UNIVERSAL table detection - look for skills-related headers
                is_skills_table = (
                    any('skill' in cell.text.lower() for cell in header_cells) or
                    any('required' in cell.text.lower() or 'desired' in cell.text.lower() for cell in header_cells) or
                    any('experience' in cell.text.lower() for cell in header_cells) or
                    len(header_cells) >= 3  # Most skills tables have at least 3 columns
                )
                
                if is_skills_table:
                    print(f"  Found skills table at index {table_idx} with {len(header_cells)} columns")
                    
                    # Remove existing data rows (keep header row only)
                    while len(table.rows) > 1:
                        table._tbl.remove(table.rows[1]._tr)
                    
                    # Add new skills from skills_data - UNIVERSAL FORMAT
                    skills_added_count = 0
                    for skill in skills_data:
                        if not skill or not skill.get('skill') or skill['skill'] == 'N/A':
                            continue
                        
                        # Add a new row for each skill
                        row_cells = table.add_row().cells
                        
                        # SAFE column access - handle different column counts
                        num_columns = len(row_cells)
                        
                        # Map data to appropriate columns based on table structure
                        if num_columns >= 1:
                            row_cells[0].text = skill.get('skill', '') or ""
                        
                        if num_columns >= 2:
                            # Map 'Required'/'Desired' based on skill type
                            skill_type = skill.get('type', '')
                            if 'required' in skill_type.lower():
                                row_cells[1].text = 'Required'
                            elif 'desired' in skill_type.lower() or 'highly' in skill_type.lower():
                                row_cells[1].text = 'Highly desired'
                            else:
                                row_cells[1].text = skill_type or ""
                        
                        if num_columns >= 3:
                            experience = skill.get('experience', '')
                            # Extract just the numeric part (e.g., "5 years" → "5")
                            years_match = re.search(r'(\d+)\s*(year|yr|years)?', experience, re.IGNORECASE)
                            years_value = years_match.group(1) if years_match else skill.get('years', '0')
                            row_cells[2].text = years_value
                        
                        # Handle additional columns if they exist
                        if num_columns >= 4:
                            # For "Years Used" column (same as years of experience)
                            row_cells[3].text = years_value if 'years_value' in locals() else skill.get('years', '0')
                        
                        if num_columns >= 5:
                            # For "Last Used" column (can be left empty or use current year)
                            row_cells[4].text = "Current"  # or leave empty ""
                        
                        skills_added_count += 1
                    
                    print(f"    Added {skills_added_count} skills to the {state_abbr} RTR table.")
                    skills_updated_in_rtr = True
                    set_table_borders(table)
                    break

        if not skills_updated_in_rtr:
            print(f"  No skills table found within {state_abbr} RTR document.")
    else:
        print(f"  No skills data available for {state_abbr} document.")
    
    # ENHANCED: Handle combined document naming (like RTR_SM_Idaho)
    output_filename = ""
    template_basename = os.path.basename(template_path).upper()
    
    # Check if this is a combined document (contains both RTR and SM in name)
    if 'RTR' in template_basename and 'SM' in template_basename:
        output_filename = f"RTR_SM_{state_abbr}_{req_id}.docx"
        print(f"  📄 Detected combined document, using naming: {output_filename}")
    else:
        output_filename = f"RTR_{state_abbr}_{req_id}.docx"
    
    output_path = os.path.join(os.path.dirname(output_path), output_filename)
    
    # Save the document
    try:
        doc.save(output_path)
        print(f"✅ Saved RTR document: {os.path.basename(output_path)}")
        return output_path
    except Exception as e:
        print(f"❌ Error saving RTR document: {e}")
        return None

def update_sm_document(skills_data, state_abbr, documents_dir, output_path, requisition_content=None):
    """
    Update the SM document with skills data for the specific state
    ENHANCED: Handles combined documents like RTR_SM_Idaho.docx
    """
    # If state_abbr is DEFAULT or unclear, try to extract from requisition content
    if state_abbr == 'DEFAULT' and requisition_content:
        state_abbr = extract_state_from_job_id(requisition_content)
        print(f"  Extracted state from content: {state_abbr}")
    
    # Get the correct template based on state - support both .doc and .docx
    state_info = STATE_MAPPING.get(state_abbr, STATE_MAPPING['DEFAULT'])
    
    # ENHANCED: Check if this state uses a combined document
    sm_template_name = state_info['sm_template']
    rtr_template_name = state_info['rtr_template']
    
    # If both templates point to the same file, it's a combined document
    is_combined_document = (sm_template_name == rtr_template_name)
    
    if is_combined_document:
        print(f"  🔄 Detected combined document state: {state_abbr}")
        print(f"  Using combined template: {sm_template_name}")
        
        # For combined documents, we need to extract req_id from output_path
        req_id_match = re.search(r'(\d+)\.docx$', output_path)
        req_id = req_id_match.group(1) if req_id_match else "unknown"
        
        # The SM document is the same as the RTR document for combined states
        combined_filename = f"RTR_SM_{state_abbr}_{req_id}.docx"
        combined_path = os.path.join(os.path.dirname(output_path), combined_filename)
        
        # Check if the combined document was already created by update_rtr_document
        if os.path.exists(combined_path):
            print(f"  ✅ Using existing combined document for SM: {combined_filename}")
            return combined_path
        else:
            print(f"  ❌ Combined document not found for SM: {combined_filename}")
            return None
    
    # Existing separate document logic for non-combined states
    # First try .docx version
    template_path_docx = os.path.join(documents_dir, state_info['sm_template'])
    template_path_doc = os.path.join(documents_dir, state_info['sm_template'].replace('.docx', '.doc'))
    
    # Check which template exists
    if os.path.exists(template_path_docx):
        template_path = template_path_docx
    elif os.path.exists(template_path_doc):
        template_path = template_path_doc
        print(f"  Using .doc template instead of .docx")
    else:
        print(f"❌ STATE TEMPLATE NOT FOUND: {state_info['sm_template']} for {state_abbr}")
        
        # Try to find any SM template that might work for this state (both .doc and .docx)
        all_sm_templates = []
        for ext in ['.docx', '.doc']:
            all_sm_templates.extend([f for f in os.listdir(documents_dir) 
                                   if f.startswith('SM_') and f.endswith(ext)])
        
        # Prioritize templates that match the state abbreviation
        state_specific_templates = [f for f in all_sm_templates if state_abbr in f.upper()]
        
        if state_specific_templates:
            template_path = os.path.join(documents_dir, state_specific_templates[0])
            print(f"   Using state-specific template: {state_specific_templates[0]}")
        elif all_sm_templates:
            print(f"   Available SM templates: {all_sm_templates}")
            template_path = os.path.join(documents_dir, all_sm_templates[0])
            print(f"   Using alternative template: {all_sm_templates[0]}")
        else:
            # Fallback to default (try both extensions)
            default_docx = os.path.join(documents_dir, STATE_MAPPING['DEFAULT']['sm_template'])
            default_doc = os.path.join(documents_dir, STATE_MAPPING['DEFAULT']['sm_template'].replace('.docx', '.doc'))
            
            if os.path.exists(default_docx):
                template_path = default_docx
            elif os.path.exists(default_doc):
                template_path = default_doc
            else:
                print(f"❌ ERROR: No default template found either!")
                return None
    
    print(f"✅ Using SM template: {os.path.basename(template_path)} for state {state_abbr}")
    
    # Load the document
    try:
        doc = Document(template_path)
    except Exception as e:
        print(f"❌ Error loading template {template_path}: {e}")
        return None
    
    # Initialize num_columns with a default value to prevent UnboundLocalError
    num_columns = 0
    skills_added = 0
    
    # Update skills table - SAFELY handle different table structures
    if doc.tables:
        table = doc.tables[0]
        
        # Get number of columns from the first row
        if len(table.rows) > 0:
            num_columns = len(table.rows[0].cells)
        
        # Remove existing data rows (keep header row)
        for i in range(len(table.rows) - 1, 0, -1):
            table._tbl.remove(table.rows[i]._tr)
        
        # Add new skills - SAFELY handle different column counts
        for skill in skills_data:
            if not skill['skill'] or skill['skill'] == 'N/A':
                continue
                
            row_cells = table.add_row().cells
            
            # SAFE column access - check how many columns exist
            if num_columns >= 1:
                row_cells[0].text = skill['skill']
            if num_columns >= 2:
                row_cells[1].text = skill['type']
            if num_columns >= 3:
                row_cells[2].text = skill['years']
            if num_columns >= 4:
                row_cells[3].text = skill['years']
            if num_columns >= 5:
                row_cells[4].text = ""
            
            skills_added += 1
        
        print(f"   Added {skills_added} skills to SM document (table has {num_columns} columns)")
        set_table_borders(table)
    else:
        print(f"   No tables found in SM document, no skills added")
    
    # Ensure correct output filename with state code (always save as .docx)
    output_path = output_path.replace("_ID_", f"_{state_abbr}_")
    if f"_{state_abbr}_" not in output_path:
        base_name = os.path.basename(output_path)
        if "SM_" in base_name:
            output_path = os.path.join(os.path.dirname(output_path), f"SM_{state_abbr}_{req_id}.docx")
        else:
            output_path = os.path.join(os.path.dirname(output_path), f"SM_{state_abbr}_{req_id}.docx")
    
    # Save the document
    try:
        doc.save(output_path)
        print(f"✅ Saved SM document: {os.path.basename(output_path)}")
        return output_path
    except Exception as e:
        print(f"❌ Error saving SM document: {e}")
        return None
    
def extract_all_skills_from_requisition(content):
    """
    Extract ALL skills from requisition content text
    Returns list of skill strings
    """
    skills = []
    
    if not content:
        return skills
    
    # Try multiple patterns to find the skills section
    patterns = [
        r'^skills:\s*\n(.*?)(?=\n\n|\n===|\nDescription:|\nJob ID:|\nLocation:|\nDuration:|\nPositions:|$)',
        r'^SKILLS:\s*\n(.*?)(?=\n\n|\n===|\nDescription:|\nJob ID:|\nLocation:|\nDuration:|\nPositions:|$)',
        r'skills:\s*\n(.*?)(?=\n\n|\n===|\nDescription:|\nJob ID:|\nLocation:|\nDuration:|\nPositions:|$)',
    ]
    
    skills_text = ""
    for pattern in patterns:
        match = re.search(pattern, content, re.DOTALL | re.MULTILINE | re.IGNORECASE)
        if match:
            skills_text = match.group(1).strip()
            break
    
    if not skills_text:
        # If no skills section found, return empty
        return skills
    
    # Parse skills from the text
    lines = skills_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Check for bullet points (•, -, *, etc.)
        if line.startswith(('•', '-', '*', '○', '∙', '·')):
            # Remove bullet point
            clean_line = re.sub(r'^[•\-*○∙·\s]+', '', line)
            clean_line = clean_line.strip()
            if clean_line:
                skills.append(clean_line)
        # Check for numbered items (1., 2., etc.)
        elif re.match(r'^\d+[\.\)]\s+', line):
            # Remove number
            clean_line = re.sub(r'^\d+[\.\)]\s+', '', line)
            clean_line = clean_line.strip()
            if clean_line:
                skills.append(clean_line)
        # Check if line looks like a skill (contains experience/years keywords)
        elif any(keyword in line.lower() for keyword in ['experience', 'knowledge', 'familiarity', 'ability', 'skill', 'proficiency', 'certification', 'years']):
            skills.append(line)
    
    return skills

def clear_updated_documents():
    """Clear the updated_documents directory before processing"""
    updated_docs_dir = "updated_documents"
    
    if not os.path.exists(updated_docs_dir):
        os.makedirs(updated_docs_dir)
        print(f"Created directory: {updated_docs_dir}")
    else:
        # Remove all files in the directory
        for file in os.listdir(updated_docs_dir):
            file_path = os.path.join(updated_docs_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    print(f"Removed previous file: {file}")
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
        print(f"Cleared output directory: {updated_docs_dir}")

def extract_title_from_requisition(content):
    """
    Extract the job title from requisition content and clean it up
    """
    # First, try to find the specific "Title/Role:" pattern
    title_match = re.search(r'Title/Role:\s*(.+)$', content, re.MULTILINE | re.IGNORECASE)
    if title_match:
        title = title_match.group(1).strip()
        # Clean up the title - remove any prefix like "NC FAST Requisition Class: DEV : "
        clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', title, flags=re.IGNORECASE)
        clean_title = clean_title.strip()
        return clean_title
    
    # Fallback: Look for patterns that might indicate a title
    lines = content.split('\n')
    for line in lines:
        line = line.strip()
        # Skip empty lines and section headers
        if not line or line.startswith('==') or len(line) > 100:
            continue
            
        # Common title indicators (case insensitive)
        if any(keyword in line.lower() for keyword in 
               ['engineer', 'developer', 'analyst', 'manager', 'specialist', 
                'consultant', 'architect', 'administrator', 'designer']):
            # Clean up the title - remove any prefix like "NC FAST Requisition Class: DEV : "
            clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', line, flags=re.IGNORECASE)
            clean_title = clean_title.strip()
            return clean_title
                
    # Final fallback: return the first meaningful line
    for line in lines:
        if line.strip() and not line.startswith('=='):
            clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', line.strip(), flags=re.IGNORECASE)
            return clean_title.strip()
    
    return "Position"

def process_requisition_files():
    """
    Main function to process all requisition files and update state-specific documents
    This runs BEFORE LLM processing
    """
    output_dir = "requisition_outputs"
    documents_dir = "Documents"  # Folder containing state-specific templates
    updated_docs_dir = "updated_documents"
    
    # First, show what templates are available
    print("\n=== AVAILABLE TEMPLATES IN DOCUMENTS FOLDER ===")
    template_files = []
    for file in os.listdir(documents_dir):
        if file.endswith('.docx'):
            template_files.append(file)
            print(f"  - {file}")
    
    if not template_files:
        print("  ❌ No template files found in Documents folder!")
    
    # Clear previous documents before processing
    clear_updated_documents()
    
    # Process each requisition file
    processed_count = 0
    for filename in os.listdir(output_dir):
        if filename.startswith("requisition_") and filename.endswith("_complete.txt"):
            filepath = os.path.join(output_dir, filename)
            
            # Extract requisition ID from filename
            match = re.search(r'requisition_(\d+)_complete\.txt', filename)
            if not match:
                print(f"Skipping file with unexpected format: {filename}")
                continue
                
            req_id = match.group(1)
            
            # Read the file content
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception as e:
                print(f"Error reading file {filename}: {e}")
                continue
            
            # Extract state abbreviation from content
            state_abbr = extract_state_from_job_id(content)
            print(f"\n{'='*60}")
            print(f"📋 PROCESSING REQUISITION {req_id} FOR STATE: {state_abbr}")
            print(f"{'='*60}")
            
            # Split into main content and skills table
            if "=== SKILLS TABLE ===" in content:
                parts = content.split("=== SKILLS TABLE ===")
                main_content = parts[0].strip()
                skills_content = parts[1].strip()
                # Remove questions section if present
                if "=== QUESTIONS ===" in skills_content:
                    skills_content = skills_content.split("=== QUESTIONS ===")[0].strip()
            else:
                main_content = content.strip()
                skills_content = ""
                print(f"  ⚠️ No skills table found in file")
            
            # === MODIFIED: Extract title from main content (previously called extract_title_from_requisition)
            title = "Position"  # Default fallback
            # First, try to find the specific "Title/Role:" pattern
            title_match = re.search(r'Title/Role:\s*(.+)$', main_content, re.MULTILINE | re.IGNORECASE)
            if title_match:
                title = title_match.group(1).strip()
                # Clean up the title - remove any prefix like "NC FAST Requisition Class: DEV : "
                clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', title, flags=re.IGNORECASE)
                clean_title = clean_title.strip()
                title = clean_title
            else:
                # Fallback: Look for patterns that might indicate a title
                lines = main_content.split('\n')
                for line in lines:
                    line = line.strip()
                    # Skip empty lines and section headers
                    if not line or line.startswith('==') or len(line) > 100:
                        continue
                        
                    # Common title indicators (case insensitive)
                    if any(keyword in line.lower() for keyword in 
                           ['engineer', 'developer', 'analyst', 'manager', 'specialist', 
                            'consultant', 'architect', 'administrator', 'designer']):
                        # Clean up the title - remove any prefix like "NC FAST Requisition Class: DEV : "
                        clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', line, flags=re.IGNORECASE)
                        clean_title = clean_title.strip()
                        title = clean_title
                        break
                # Final fallback: return the first meaningful line
                for line in lines:
                    if line.strip() and not line.startswith('=='):
                        clean_title = re.sub(r'^.*?(?:requisition class|nc fast)[^:]*:\s*', '', line.strip(), flags=re.IGNORECASE)
                        title = clean_title.strip()
                        break
            
            print(f"📝 Extracted title: {title}")
            
                        # Extract skills from content
            print(f"🔧 Extracting skills from content...")
            skills_data = []

            # FIRST: Check for regex-extracted format (SKILLS TABLE: section)
            if "SKILLS TABLE:" in content:
                print(f"  📊 Found 'SKILLS TABLE:' section (regex-extracted format)")
                
                # Extract everything from "SKILLS TABLE:" to the next major section
                start_idx = content.find("SKILLS TABLE:")
                if start_idx != -1:
                    # Get the substring starting from SKILLS TABLE:
                    from_skills = content[start_idx:]
                    
                    # Find the next major section
                    next_sections = [
                        "Description:", "SHORT DESCRIPTION:", "COMPLETE DESCRIPTION:", 
                        "Job ID:", "Location:", "Duration:", "Positions:",
                        "\n\n\n", "\n\n"
                    ]
                    
                    end_idx = len(from_skills)
                    for section in next_sections:
                        idx = from_skills.find(section)
                        if idx != -1 and idx < end_idx:
                            end_idx = idx
                    
                    # Extract skills text
                    skills_text = from_skills[:end_idx].strip()
                    
                    # Parse using our new function
                    skills_data = parse_regex_extracted_skills(skills_text)
                    
                    if skills_data:
                        print(f"  ✅ Successfully extracted {len(skills_data)} skills from regex format")

            # SECOND: If no skills found, try original markdown table format
            if not skills_data and "=== SKILLS TABLE ===" in content:
                print(f"  📊 Found '=== SKILLS TABLE ===' (original scraped format)")
                parts = content.split("=== SKILLS TABLE ===")
                if len(parts) > 1:
                    skills_content = parts[1].split("\n\n")[0].strip()
                    skills_data = parse_skills_table(skills_content)
                    print(f"  ✅ Parsed {len(skills_data)} skills from markdown table")

            # THIRD: If still no skills, try LLM format
            if not skills_data and "skills:" in content.lower():
                print(f"  🤖 Found 'skills:' section (LLM format)")
                # Try to extract skills in LLM format
                lines = content.split('\n')
                in_skills = False
                skills_lines = []
                
                for line in lines:
                    if line.strip().lower() == 'skills:':
                        in_skills = True
                        continue
                    elif in_skills and line.strip():
                        if any(line.strip().startswith(section) for section in 
                              ['Description:', 'Job ID:', 'Location:', 'Duration:', 'Positions:']):
                            break
                        skills_lines.append(line.strip())
                    elif in_skills and not line.strip():
                        break
                
                if skills_lines:
                    skills_text = '\n'.join(skills_lines)
                    skills_data = parse_regex_extracted_skills(skills_text)
                    print(f"  ✅ Parsed {len(skills_data)} skills from LLM format")

            # FOURTH: If still no skills, create placeholder
            if not skills_data:
                print(f"  ⚠️ No skills found, creating placeholder")
                skills_data = [{
                    'skill': f"{title} Skills",
                    'type': 'Required',
                    'experience': '',
                    'years': '5'
                }]
                print(f"  ⚠️ Created placeholder skill from title")

            # DEBUG: Show final skills
            print(f"  📋 Final skills count: {len(skills_data)}")
            if skills_data:
                print(f"  📋 Skills to be used in SM document:")
                for i, skill in enumerate(skills_data[:5]):  # Show first 5
                    print(f"    {i+1}. {skill.get('skill', 'N/A')}")
                if len(skills_data) > 5:
                    print(f"    ... and {len(skills_data) - 5} more")
            else:
                print(f"  ❌ No skills data available!")
                # Create at least one skill from the title
                skills_data = [{
                    'skill': f"{title} Skills",
                    'type': 'Required',
                    'experience': '',
                    'years': '5'
                }]
                print(f"  ⚠️ Created placeholder skill from title")
            
            # Update state-specific documents
            rtr_output = os.path.join(updated_docs_dir, f"RTR_{state_abbr}_{req_id}.docx")
            sm_output = os.path.join(updated_docs_dir, f"SM_{state_abbr}_{req_id}.docx")
            
            print(f"\n--- UPDATING RTR DOCUMENT ---")
            actual_rtr_path = update_rtr_document(req_id, title, state_abbr, documents_dir, 
                                                  rtr_output, requisition_content=content, 
                                                  skills_data=skills_data)
            
            print(f"\n--- UPDATING SM DOCUMENT ---")
            # Pass ALL parameters including the full content
            actual_sm_path = update_sm_document(
                skills_data, 
                state_abbr, 
                documents_dir, 
                sm_output, 
                requisition_content=content  # Pass the full content
            )
            
            # Verify the documents were created successfully
            if actual_rtr_path and actual_sm_path:
                # Check if files were saved with correct state codes
                rtr_filename = os.path.basename(actual_rtr_path)
                sm_filename = os.path.basename(actual_sm_path)
                
                if f"_{state_abbr}_" in rtr_filename and f"_{state_abbr}_" in sm_filename:
                    print(f"✅ SUCCESS: Documents created with correct state codes")
                    print(f"   📄 RTR: {rtr_filename}")
                    print(f"   📄 SM: {sm_filename}")
                else:
                    print(f"⚠️  WARNING: Documents may have incorrect naming:")
                    print(f"   📄 RTR: {rtr_filename}")
                    print(f"   📄 SM: {sm_filename}")
                
                # Verify files actually exist
                rtr_exists = os.path.exists(actual_rtr_path)
                sm_exists = os.path.exists(actual_sm_path)
                
                if rtr_exists and sm_exists:
                    print(f"✅ Files verified on disk")
                    processed_count += 1
                else:
                    print(f"❌ ERROR: Files not found on disk:")
                    print(f"   RTR exists: {rtr_exists}")
                    print(f"   SM exists: {sm_exists}")
            else:
                print(f"❌ FAILED: Could not create documents for requisition {req_id}")
                if not actual_rtr_path:
                    print(f"   RTR document creation failed")
                if not actual_sm_path:
                    print(f"   SM document creation failed")
            
            print(f"\n{'='*60}")
    
    print(f"\n=== PROCESSING COMPLETE ===")
    print(f"✅ Successfully processed {processed_count} requisitions")
    
    # Show all created documents
    print(f"\n=== CREATED DOCUMENTS IN {updated_docs_dir} ===")
    sm_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('SM_') and f.endswith('.docx')]
    rtr_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('RTR_') and f.endswith('.docx')]
    
    print(f"📄 SM Documents: {len(sm_files)}")
    for sm_file in sm_files:
        state_from_filename = re.search(r'SM_([A-Z]{2})_', sm_file)
        state_code = state_from_filename.group(1) if state_from_filename else "UNKNOWN"
        print(f"  - {sm_file} (State: {state_code})")
    
    print(f"📄 RTR Documents: {len(rtr_files)}")
    for rtr_file in rtr_files:
        state_from_filename = re.search(r'RTR_([A-Z]{2})_', rtr_file)
        state_code = state_from_filename.group(1) if state_from_filename else "UNKNOWN"
        print(f"  - {rtr_file} (State: {state_code})")
    
    return processed_count

def extract_questions_section(driver, max_retries=3):
    """
    Extract the Questions section from the Vector VMS requisition page.
    Returns a formatted string with all questions or empty string if not found.
    """
    for attempt in range(max_retries):
        try:
            print(f"  Attempt {attempt + 1} to extract Questions section...")
            
            # Switch to iframe if present (this is often needed for VMS content)
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    driver.switch_to.frame(iframes[0])
                    print("  Switched to iframe for content")
            except:
                print("  No iframe detected or error switching")
                # Continue anyway - the content might be in the main frame

            # Look for the Questions section by its ID or header text
            questions_content = ""
            
            # Method 1: Look for the specific panel ID
            try:
                questions_panel = driver.find_element(By.ID, "ContentPH_pnlQuestions")
                print("  Found Questions panel by ID")
                questions_content = questions_panel.text
            except:
                # Method 2: Look for the header text "Questions"
                try:
                    questions_header = driver.find_element(By.XPATH, "//*[contains(text(), 'Questions')]")
                    print("  Found Questions header text")
                    # Get the parent container of the header and extract its text
                    questions_container = questions_header.find_element(By.XPATH, "./ancestor::div[contains(@class, 'x-panel')]")
                    questions_content = questions_container.text
                except:
                    # Method 3: Look for any element containing question text
                    try:
                        question_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Question') or contains(text(), 'question')]")
                        if question_elements:
                            print(f"  Found {len(question_elements)} question elements")
                            questions_content = "\n".join([elem.text for elem in question_elements if elem.text.strip()])
                    except Exception as e:
                        print(f"  Error finding question elements: {e}")
            
            # If we found content, parse it for actual questions
            if questions_content:
                print(f"  Raw questions content found: {len(questions_content)} characters")
                
                # Parse the content to extract just the question text
                lines = questions_content.split('\n')
                questions = []
                current_question = ""
                in_question = False
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                        
                    # Look for question indicators
                    if line.startswith('Question') or line.startswith('Q:') or (len(line) > 20 and '?' in line and not line.startswith('Description')):
                        if current_question and current_question not in questions:
                            questions.append(current_question)
                        current_question = line
                        in_question = True
                    elif in_question and line:
                        # Continue building the current question
                        current_question += " " + line
                
                # Add the last question if exists
                if current_question and current_question not in questions:
                    questions.append(current_question)
                
                # Filter out non-question lines
                filtered_questions = []
                for q in questions:
                    if ('question' in q.lower() or '?' in q) and len(q) > 10:
                        # Clean up the question text
                        q = re.sub(r'^Question\s*\d+[:.\s]*', '', q, flags=re.IGNORECASE)
                        q = q.strip()
                        if q:
                            filtered_questions.append(q)
                
                # Format the questions output
                if filtered_questions:
                    formatted_questions = "=== QUESTIONS ===\n"
                    for i, question in enumerate(filtered_questions, 1):
                        formatted_questions += f"Q{i}: {question}\n"
                    
                    print(f"  Successfully extracted {len(filtered_questions)} questions")
                    # Switch back to default content before returning
                    try:
                        driver.switch_to.default_content()
                    except:
                        pass
                    return formatted_questions
            
            print("  No questions found in the section")
            # Switch back to default content
            try:
                driver.switch_to.default_content()
            except:
                pass
            return ""

        except Exception as e:
            print(f"  Attempt {attempt + 1} failed: {str(e)}")
            # Switch back to default content on error
            try:
                driver.switch_to.default_content()
            except:
                pass
            
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                print("  Max retries reached for Questions extraction")
                return ""

    return ""

def process_single_requisition(driver, url):
    """Process a single requisition using the credential fallback logic"""
    print(f"\n📋 Processing requisition: {url}")
    
    # Try each credential set in order
    for i, credentials in enumerate(CREDENTIAL_SETS):
        print(f"\n🔑 Trying credential set {i+1}/{len(CREDENTIAL_SETS)}")
        
        # Try to login with current credentials
        login_success = login_to_vectorvms(driver, url, credentials)
        
        if login_success:
            try:
                # Ensure we're in the default content context first
                driver.switch_to.default_content()
                time.sleep(3)
                
                print("  📄 Extracting main content...")
                # Extract content
                content = extract_complete_page_content(driver)
                
                print("  🔧 Extracting skills table...")
                # Extract skills table WITH RETRIES
                skills_table = ""
                skills_attempts = 2
                for attempt in range(skills_attempts):
                    print(f"    Skills extraction attempt {attempt + 1}/{skills_attempts}")
                    skills_table = extract_skills_table(driver)
                    if skills_table and "Technical Skills" not in skills_table:
                        print(f"    ✅ Got non-placeholder skills table")
                        break
                    elif attempt < skills_attempts - 1:
                        print(f"    ⏳ Retrying skills extraction...")
                        time.sleep(2)
                        driver.refresh()
                        time.sleep(3)
                
                # Ensure we're back in default context
                driver.switch_to.default_content()
                
                print("  ❓ Extracting questions section...")
                # Extract questions section
                questions_section = extract_questions_section(driver)
                
                # Save results
                req_id = re.search(r'reqID=(\d+)', url).group(1)
                output_dir = "requisition_outputs"
                output_file = os.path.join(output_dir, f"requisition_{req_id}_complete.txt")
                
                print(f"  💾 Saving to file: {output_file}")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                    if skills_table:
                        f.write("\n\n=== SKILLS TABLE ===\n")
                        f.write(skills_table)
                        print(f"    ✅ Saved skills table with {skills_table.count('|')//3} skills")
                    # Add questions section after skills
                    if questions_section:
                        f.write("\n\n")
                        f.write(questions_section)
                
                print(f"✅ Saved results using org_key: {credentials['org_key']}")
                
                # Logout after successful extraction
                logout_from_vectorvms(driver)
                
                return True, credentials['org_key']
                
            except Exception as e:
                print(f"❌ Error extracting data: {str(e)}")
                import traceback
                traceback.print_exc()
                # Ensure we switch back to default content on error
                try:
                    driver.switch_to.default_content()
                except:
                    pass
                # Try to logout even if extraction failed
                try:
                    logout_from_vectorvms(driver)
                except:
                    pass
                continue
        else:
            print(f"❌ Login failed with org_key: {credentials['org_key']}")
            # Clear cookies and try next credentials
            driver.delete_all_cookies()
            driver.refresh()
            time.sleep(2)
    
    print(f"❌ All credential sets failed for URL: {url}")
    return False, None

def extract_questions_section(driver, max_retries=3):
    """
    Extract the Questions section from the Vector VMS requisition page.
    Returns a formatted string with all questions or empty string if not found.
    """
    for attempt in range(max_retries):
        try:
            print(f"  Attempt {attempt + 1} to extract Questions section...")
            
            # Switch to iframe if present (this is often needed for VMS content)
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    driver.switch_to.frame(iframes[0])
                    print("  Switched to iframe for content")
            except:
                print("  No iframe detected or error switching")
                # Continue anyway - the content might be in the main frame

            # Look for the Questions section by its ID or header text
            questions_content = ""
            
            # Method 1: Look for the specific panel ID
            try:
                questions_panel = driver.find_element(By.ID, "ContentPH_pnlQuestions")
                print("  Found Questions panel by ID")
                questions_content = questions_panel.text
            except:
                # Method 2: Look for the header text "Questions"
                try:
                    questions_header = driver.find_element(By.XPATH, "//*[contains(text(), 'Questions')]")
                    print("  Found Questions header text")
                    # Get the parent container of the header and extract its text
                    questions_container = questions_header.find_element(By.XPATH, "./ancestor::div[contains(@class, 'x-panel')]")
                    questions_content = questions_container.text
                except:
                    # Method 3: Look for any element containing question text
                    try:
                        question_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Question') or contains(text(), 'question')]")
                        if question_elements:
                            print(f"  Found {len(question_elements)} question elements")
                            questions_content = "\n".join([elem.text for elem in question_elements if elem.text.strip()])
                    except Exception as e:
                        print(f"  Error finding question elements: {e}")
            
            # If we found content, parse it for actual questions
            if questions_content:
                print(f"  Raw questions content found: {len(questions_content)} characters")
                
                # Parse the content to extract just the question text
                lines = questions_content.split('\n')
                questions = []
                current_question = ""
                in_question = False
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                        
                    # Look for question indicators
                    if line.startswith('Question') or line.startswith('Q:') or (len(line) > 20 and '?' in line and not line.startswith('Description')):
                        if current_question and current_question not in questions:
                            questions.append(current_question)
                        current_question = line
                        in_question = True
                    elif in_question and line:
                        # Continue building the current question
                        current_question += " " + line
                
                # Add the last question if exists
                if current_question and current_question not in questions:
                    questions.append(current_question)
                
                # Filter out non-question lines
                filtered_questions = []
                for q in questions:
                    if ('question' in q.lower() or '?' in q) and len(q) > 10:
                        # Clean up the question text
                        q = re.sub(r'^Question\s*\d+[:.\s]*', '', q, flags=re.IGNORECASE)
                        q = q.strip()
                        if q:
                            filtered_questions.append(q)
                
                # Format the questions output
                if filtered_questions:
                    formatted_questions = "=== QUESTIONS ===\n"
                    for i, question in enumerate(filtered_questions, 1):
                        formatted_questions += f"Q{i}: {question}\n"
                    
                    print(f"  Successfully extracted {len(filtered_questions)} questions")
                    # Switch back to default content before returning
                    try:
                        driver.switch_to.default_content()
                    except:
                        pass
                    return formatted_questions
            
            print("  No questions found in the section")
            # Switch back to default content
            try:
                driver.switch_to.default_content()
            except:
                pass
            return ""

        except Exception as e:
            print(f"  Attempt {attempt + 1} failed: {str(e)}")
            # Switch back to default content on error
            try:
                driver.switch_to.default_content()
            except:
                pass
            
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                print("  Max retries reached for Questions extraction")
                return ""

    return ""

def open_all_requisitions_in_new_tabs(driver, urls):
    """Open all requisition URLs in new tabs with credential fallback logic"""
    print(f"\nOpening {len(urls)} requisitions in new tabs...")
    
    # Store the original window handle
    original_window = driver.current_window_handle
    
    # Track credential performance
    credential_stats = {cred['org_key']: {'success': 0, 'attempts': 0} for cred in CREDENTIAL_SETS}
    
    # Process each URL
    for i, url in enumerate(urls):
        try:
            # Open new tab using JavaScript
            driver.execute_script("window.open('');")
            
            # Switch to the new tab
            all_windows = driver.window_handles
            driver.switch_to.window(all_windows[-1])
            
            # Process the requisition with credential fallback logic
            success, used_org_key = process_single_requisition(driver, url)
            
            # Update statistics
            if success:
                credential_stats[used_org_key]['success'] += 1
            for cred in CREDENTIAL_SETS:
                credential_stats[cred['org_key']]['attempts'] += 1
            
            if success:
                print(f"✓ Successfully processed tab {i+1}/{len(urls)}: {url}")
            else:
                print(f"✗ Failed to process tab {i+1}/{len(urls)}: {url}")
            
            # Close the tab and switch back to the original window
            driver.close()
            driver.switch_to.window(original_window)
            
            # Brief pause between URLs
            time.sleep(2)
            
        except Exception as e:
            print(f"Error processing URL {url}: {str(e)}")
            # Try to recover by switching back to the original window
            try:
                driver.switch_to.window(original_window)
            except:
                pass
            continue
    
    # Print credential performance summary
    print("\n" + "="*60)
    print("CREDENTIAL PERFORMANCE SUMMARY:")
    print("="*60)
    for org_key, stats in credential_stats.items():
        success_rate = (stats['success'] / stats['attempts'] * 100) if stats['attempts'] > 0 else 0
        print(f"Org Key {org_key}: {stats['success']} successes out of {stats['attempts']} attempts ({success_rate:.1f}% success rate)")



# ===== REGEX EXTRACTION FUNCTIONS =====
def extract_deadline_date(text):
    """
    Extract deadline date - ONLY look for "No New Submittals After"
    If not present, calculate 4 BUSINESS DAYS from today
    Returns MMDD format string
    """
    from datetime import datetime
    import re
    
    print(f"  🔍 Looking for 'No New Submittals After' date...")
    
    # ONLY look for "No New Submittals After" patterns
    patterns = [
        # Pattern 1: "No New Submittals After: 09-16-2024"
        r'No New Submittals After:\s*(\d{1,2})-(\d{1,2})-(\d{4})',
        # Pattern 2: "No New Submittals After: 09/16/2024"
        r'No New Submittals After:\s*(\d{1,2})/(\d{1,2})/(\d{4})',
        # Pattern 3: "No New Submittals After: September 16, 2024"
        r'No New Submittals After:\s*([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})',
        # Pattern 4: "No New Submittals After: 09/16" (NEW - without year)
        r'No New Submittals After:\s*(\d{1,2})/(\d{1,2})\b',
        # Pattern 5: "No New Submittals After: 09-16" (NEW - without year)
        r'No New Submittals After:\s*(\d{1,2})-(\d{1,2})\b',
    ]
    
    # ==============================================
    # LOGIC: Check if phrase is PRESENT or NOT
    # ==============================================
    
    # Check if "No New Submittals After" exists anywhere in text
    if 'No New Submittals After' not in text:
        print(f"  ⚠️  'No New Submittals After' NOT PRESENT in text")
        print(f"  📅 Going to 4 BUSINESS DAYS calculation (Virginia logic)")
        return calculate_virginia_deadline()
    
    # If we get here, "No New Submittals After" IS PRESENT in text
    print(f"  ✅ 'No New Submittals After' IS PRESENT in text")
    
    # Try to extract the date using patterns
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            print(f"  ✅ Successfully extracted date from 'No New Submittals After'")
            try:
                # For numeric dates with year (09-16-2024 or 09/16/2024)
                if pattern in [patterns[0], patterns[1]]:
                    month = match.group(1).zfill(2)  # "9" -> "09"
                    day = match.group(2).zfill(2)    # "16" -> "16"
                    year = match.group(3)
                    print(f"  📅 Extracted deadline: {month}/{day}/{year}")
                    return f"{month}{day}"  # Return "0916"
                
                # For text month dates (September 16, 2024)
                elif pattern == patterns[2]:
                    month_name = match.group(1)
                    day = match.group(2).zfill(2)
                    year = match.group(3)
                    
                    # Convert month name to number
                    month_dict = {
                        'january': '01', 'february': '02', 'march': '03', 'april': '04',
                        'may': '05', 'june': '06', 'july': '07', 'august': '08',
                        'september': '09', 'october': '10', 'november': '11', 'december': '12',
                        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
                        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
                        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
                    }
                    
                    month_lower = month_name.lower()
                    if month_lower in month_dict:
                        month = month_dict[month_lower]
                        print(f"  📅 Extracted deadline: {month}/{day}/{year}")
                        return f"{month}{day}"
                        
                # For dates WITHOUT year (12/11 or 12-11) - NEW PATTERNS
                elif pattern in [patterns[3], patterns[4]]:
                    month = match.group(1).zfill(2)  # "12" -> "12"
                    day = match.group(2).zfill(2)    # "11" -> "11"
                    # Get current year for display
                    current_year = datetime.now().year
                    print(f"  📅 Extracted deadline without year: {month}/{day} (assuming {current_year})")
                    return f"{month}{day}"  # Return "1211"
                        
            except Exception as e:
                print(f"  ⚠️  Error extracting date: {e}")
                continue
    
    # If we get here, "No New Submittals After" is PRESENT but date couldn't be extracted
    print(f"  ⚠️  'No New Submittals After' found but COULD NOT EXTRACT DATE")
    print(f"  📅 Going to 4 BUSINESS DAYS calculation")
    return calculate_virginia_deadline()

def calculate_virginia_deadline():
    """
    Calculate deadline for Virginia locations ONLY - 4 business days from today
    Should only be called when NO date is found in the content
    """
    from datetime import datetime, timedelta

    today = datetime.now()
    days_added = 0
    target_days = 4

    print(f"  📅 Calculating Virginia deadline (4 business days from today)")
    print(f"  📅 Today: {today.strftime('%A, %Y-%m-%d')}")

    # Start from tomorrow
    current_date = today + timedelta(days=1)
    
    while days_added < target_days:
        # Monday=0, Friday=4, Saturday=5, Sunday=6
        if current_date.weekday() < 5:  # Weekday (Mon-Fri)
            days_added += 1
        # Move to next day
        current_date += timedelta(days=1)
    
    # Go back one day (we overshot by 1)
    deadline_date = current_date - timedelta(days=1)
    deadline = deadline_date.strftime("%m%d")
    
    print(f"  📅 Calculated Virginia deadline: {deadline} ({deadline_date.strftime('%A, %B %d, %Y')})")
    return deadline

# ===== REGEX EXTRACTION FUNCTIONS =====
def extract_essential_data_with_regex(raw_content):
    """
    Extract all essential information from raw scraped data using regex
    Returns cleaned, structured dictionary with proper formatting
    """
    extracted = {}
    
    # === TITLE/ROLE EXTRACTION ===
    title_match = re.search(r'Title/Role:\s*(.+)', raw_content)
    if title_match:
        extracted['title'] = title_match.group(1).strip()
        print(f"  📝 Extracted Title/Role: {extracted['title']}")
    else:
        # Try alternative patterns if "Title/Role:" not found
        alt_title_patterns = [
            r'Title:\s*(.+)',
            r'Role:\s*(.+)', 
            r'Job Title:\s*(.+)',
            r'Position Title:\s*(.+)'
        ]
        
        for pattern in alt_title_patterns:
            match = re.search(pattern, raw_content)
            if match:
                extracted['title'] = match.group(1).strip()
                print(f"  📝 Extracted Title (alternative pattern): {extracted['title']}")
                break
        
        if 'title' not in extracted:
            extracted['title'] = "Position"
            print("  ⚠️  No title found, using default: 'Position'")
    
        # === WORKSITE ADDRESS ===
    # Capture multi-line address (until next section or blank line)
    worksite_match = re.search(r'Worksite Address:\s*(.*?)(?=\n\s*(?:CAI|Max Submittals|Per Opening|Currently Engaged|Expenses Allowed|Work Arrangement|Title/Role):|\n\n|\Z)', 
                              raw_content, re.DOTALL | re.IGNORECASE)
    
    if worksite_match:
        worksite = worksite_match.group(1).strip()
        
        # If there are newlines in the address, join them with space
        if '\n' in worksite:
            # Split by lines, strip each line, then join with space
            lines = [line.strip() for line in worksite.split('\n') if line.strip()]
            worksite = ' '.join(lines)
            print(f"  🏢 Extracted multi-line Worksite Address: {worksite}")
        else:
            print(f"  🏢 Extracted Worksite Address: {worksite}")
        
        extracted['worksite_address'] = worksite
    else:
        # Fallback: try single line pattern
        worksite_match = re.search(r'Worksite Address:\s*(.+?)(?=\n|$)', raw_content)
        if worksite_match:
            worksite = worksite_match.group(1).strip()
            extracted['worksite_address'] = worksite
            print(f"  🏢 Extracted Worksite Address (single line): {worksite}")
    
    # === JOB ID & BASIC INFO ===
    job_id_match = re.search(r'Job ID:\s*([A-Z]{2}-\d+)', raw_content)
    extracted['job_id'] = job_id_match.group(1) if job_id_match else "UNKNOWN"
    
    # === MAX SUBMITTALS BY VENDOR ===
    max_submittals_match = re.search(r'Max Submittals by Vendor:\s*(\d+)', raw_content)
    extracted['max_submittals'] = max_submittals_match.group(1) if max_submittals_match else "1"
    
    # === LOCATION ===
    # Simplified version that definitely works for your format
    work_location_match = re.search(r'Work\s*Location:\s*([A-Za-z0-9\-_]+)',raw_content)
    if work_location_match:
        work_location = work_location_match.group(1).strip()
        print(f"  🏢 Work Location: {work_location}")
    else:
        work_location = ""
        print("  ⚠️  Work Location not found")
    
    # === DATES & OPENINGS - EXTRACT ALL FIELDS ===
    # No. of Openings
    openings_match = re.search(r'No\. of Openings:\s*(\d+)', raw_content)
    extracted['openings'] = openings_match.group(1) if openings_match else "1"
    
    # Total No. Filled
    filled_match = re.search(r'Total No\. Filled:\s*(\d+)', raw_content)
    extracted['filled'] = filled_match.group(1) if filled_match else "0"
    
    # Start Date
    start_match = re.search(r'Start Date:\s*(\d{1,2}/\d{1,2}/\d{4})', raw_content)
    extracted['start_date'] = start_match.group(1) if start_match else ""
    
    # End Date
    end_match = re.search(r'End Date:\s*(\d{1,2}/\d{1,2}/\d{4})', raw_content)
    extracted['end_date'] = end_match.group(1) if end_match else ""
    
    # === DEADLINE EXTRACTION WITH VIRGINIA LOGIC ===
    deadline_result = extract_deadline_date(raw_content)
    extracted['deadline'] = deadline_result  # This will be in MMDD format
    
    # === WORK ARRANGEMENT ===
    work_arrangement_match = re.search(r'Work Arrangement:\s*([^\n]+)', raw_content)
    extracted['work_arrangement'] = work_arrangement_match.group(1).strip() if work_arrangement_match else "Onsite"
    
    # === BILL RATE EXTRACTION - FROM BOTH SECTIONS SEPARATELY ===
    bill_rates = extract_bill_rates_from_all_sections(raw_content)
    extracted.update(bill_rates)
    
    # === DESCRIPTIONS - PRESERVE ORIGINAL FORMATTING ===
    # Short Description - keep original formatting
    short_desc_match = re.search(r'SHORT DESCRIPTION:\s*(.*?)(?=COMPLETE DESCRIPTION:|\n\n|\Z)', raw_content, re.DOTALL)
    if short_desc_match:
        short_desc = short_desc_match.group(1).strip()
        # Remove extra whitespace but preserve line breaks
        short_desc = re.sub(r'[ \t]+', ' ', short_desc)  # Replace multiple spaces/tabs with single space
        short_desc = re.sub(r'\n[ \t]+\n', '\n\n', short_desc)  # Preserve paragraph breaks
        extracted['short_description'] = short_desc
    
    # Complete Description - keep original formatting  
    complete_desc_match = re.search(r'COMPLETE DESCRIPTION:\s*(.*?)(?=\n\n|\Z|===)', raw_content, re.DOTALL)
    if complete_desc_match:
        complete_desc = complete_desc_match.group(1).strip()
        # Remove extra whitespace but preserve line breaks
        complete_desc = re.sub(r'[ \t]+', ' ', complete_desc)  # Replace multiple spaces/tabs with single space
        complete_desc = re.sub(r'\n[ \t]+\n', '\n\n', complete_desc)  # Preserve paragraph breaks
        # Remove manager notes if present
        complete_desc = re.sub(r'MANAGER NOTES:.*?Description:', '', complete_desc, flags=re.DOTALL)
        extracted['complete_description'] = complete_desc
    
    # === SKILLS TABLE - CONVERT TO BULLET POINT FORMAT ===
    skills_section_match = re.search(r'=== SKILLS TABLE ===(.*?)(?=\n\n|\Z|===)', raw_content, re.DOTALL)
    if skills_section_match:
        skills_table = skills_section_match.group(1).strip()
    
        bullet_points = []
        for line in skills_table.split('\n'):
            line = line.strip()
            # Process only table data rows (skip headers and separators)
            if line.startswith('|') and '---' not in line and 'Skill' not in line:
                cells = [cell.strip() for cell in line.split('|')[1:-1] if cell.strip()]
                if len(cells) >= 3:
                    bullet_points.append(f"{cells[0]}. {cells[1]} {cells[2]}")
    
        extracted['skills_table'] = "\n".join(bullet_points) if bullet_points else ""
    
    return extracted

def extract_deadline_date(text):
    """
    Extract deadline date - ONLY look for "No New Submittals After"
    If not present, calculate 4 BUSINESS DAYS from today
    Returns MMDD format string
    """
    from datetime import datetime, timedelta
    import re
    
    print(f"  🔍 Looking for 'No New Submittals After' date...")
    
    # ONLY look for "No New Submittals After" patterns
    patterns = [
        # Pattern 1: "No New Submittals After: 09-16-2024"
        r'No New Submittals After:\s*(\d{1,2})-(\d{1,2})-(\d{4})',
        # Pattern 2: "No New Submittals After: 09/16/2024"
        r'No New Submittals After:\s*(\d{1,2})/(\d{1,2})/(\d{4})',
        # Pattern 3: "No New Submittals After: September 16, 2024"
        r'No New Submittals After:\s*([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})',
    ]
    
    # ==============================================
    # LOGIC: Check if phrase is PRESENT or NOT
    # ==============================================
    
    # Check if "No New Submittals After" exists anywhere in text
    if 'No New Submittals After' not in text:
        print(f"  ⚠️  'No New Submittals After' NOT PRESENT in text")
        print(f"  📅 Going to 4 BUSINESS DAYS calculation (Virginia logic)")
        return calculate_virginia_deadline()
    
    # If we get here, "No New Submittals After" IS PRESENT in text
    print(f"  ✅ 'No New Submittals After' IS PRESENT in text")
    
    # Try to extract the date using patterns
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            print(f"  ✅ Successfully extracted date from 'No New Submittals After'")
            try:
                # For numeric dates (09-16-2024 or 09/16/2024)
                if match.group(1).isdigit():
                    month = match.group(1).zfill(2)  # "9" -> "09"
                    day = match.group(2).zfill(2)    # "16" -> "16"
                    year = match.group(3)
                    print(f"  📅 Extracted deadline: {month}/{day}/{year}")
                    return f"{month}{day}"  # Return "0916"
                
                # For text month dates (September 16, 2024)
                else:
                    month_name = match.group(1)
                    day = match.group(2).zfill(2)
                    year = match.group(3)
                    
                    # Convert month name to number
                    month_dict = {
                        'january': '01', 'february': '02', 'march': '03', 'april': '04',
                        'may': '05', 'june': '06', 'july': '07', 'august': '08',
                        'september': '09', 'october': '10', 'november': '11', 'december': '12',
                        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
                        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
                        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
                    }
                    
                    month_lower = month_name.lower()
                    if month_lower in month_dict:
                        month = month_dict[month_lower]
                        print(f"  📅 Extracted deadline: {month}/{day}/{year}")
                        return f"{month}{day}"
                        
            except Exception as e:
                print(f"  ⚠️  Error extracting date: {e}")
                continue
    
    # If we get here, "No New Submittals After" is PRESENT but date couldn't be extracted
    print(f"  ⚠️  'No New Submittals After' found but COULD NOT EXTRACT DATE")
    print(f"  📅 Going to 4 BUSINESS DAYS calculation")
    return calculate_virginia_deadline()

def calculate_virginia_deadline():
    """
    Calculate 4 BUSINESS DAYS from today
    Skips weekends (Saturday and Sunday)
    """
    from datetime import datetime, timedelta
    
    today = datetime.now()
    days_added = 0
    target_days = 4
    
    print(f"  📅 Today: {today.strftime('%A, %Y-%m-%d')}")
    print(f"  📅 Calculating 4 business days from today...")
    
    # Start from tomorrow
    current_date = today + timedelta(days=1)
    
    while days_added < target_days:
        # Monday=0, Friday=4, Saturday=5, Sunday=6
        if current_date.weekday() < 5:  # Weekday (Mon-Fri)
            days_added += 1
        # Move to next day
        current_date += timedelta(days=1)
    
    # Go back one day (we overshot by 1)
    deadline_date = current_date - timedelta(days=1)
    deadline = deadline_date.strftime("%m%d")
    
    print(f"  📅 Calculated deadline: {deadline} ({deadline_date.strftime('%A, %B %d, %Y')})")
    return deadline

def extract_bill_rates_from_all_sections(raw_content):
    """
    Extract bill rates from ALL sections
    """
    print("=" * 60)
    print("🔍 DEBUG BILL RATE EXTRACTION")
    print("=" * 60)
    
    # Print a sample to see what's in the content
    print("📋 First 1000 chars of content:")
    print(raw_content[:1000])
    print("=" * 60)
    
    bill_rates = {
        'current_budget_rate': "00",
        'questions_rate': "00", 
        'short_desc_rate': "00",
        'final_bill_rate': "00",
        'source': 'none'
    }
    
    # === 1. Check for ANY dollar amounts in ENTIRE content ===
    print("💰 Looking for ANY $ amounts in entire content...")
    all_dollar_matches = re.findall(r'\$(\d+\.?\d+)', raw_content)
    print(f"   Found: {all_dollar_matches}")
    
    # === 2. Specifically check Questions section ===
    print("\n🔍 Looking for Questions section...")
    
    # Try different patterns to find Questions
    questions_patterns = [
        r'=== QUESTIONS ===(.*?)(?===|\Z)',
        r'QUESTIONS(.*?)(?=\n\n|\Z)',
        r'Q1:(.*?)(?=Q\d+:|$)',
    ]
    
    questions_content = ""
    for pattern in questions_patterns:
        match = re.search(pattern, raw_content, re.DOTALL | re.IGNORECASE)
        if match:
            questions_content = match.group(1)
            print(f"✅ Found Questions section with pattern: {pattern[:30]}...")
            print(f"   Questions content length: {len(questions_content)} chars")
            print(f"   First 200 chars: {questions_content[:200]}")
            break
    
    if questions_content:
        print("\n💰 Searching for $ in Questions...")
        question_dollar_matches = re.findall(r'\$(\d+\.?\d+)', questions_content)
        print(f"   Found in Questions: {question_dollar_matches}")
        
        # Filter reasonable rates
        valid_rates = []
        for rate in question_dollar_matches:
            try:
                rate_num = float(rate)
                if 10 <= rate_num <= 200:
                    valid_rates.append(rate_num)
                    print(f"   ✅ Valid rate: ${rate_num}")
            except:
                continue
        
        if valid_rates:
            highest_rate = max(valid_rates)
            bill_rates['questions_rate'] = f"{highest_rate:.2f}"
            print(f"🎯 SET Questions rate to: ${bill_rates['questions_rate']}")
    
    # === 3. Check Current Budget ===
    print("\n💰 Looking for Current Budget $ amounts...")
    # Look for patterns like: $73.79 USD
    usd_matches = re.findall(r'\$(\d+\.?\d+)\s*USD', raw_content)
    print(f"   USD matches: {usd_matches}")
    
    valid_rates = []
    for rate in usd_matches:
        try:
            rate_num = float(rate)
            if 10 <= rate_num <= 200:
                valid_rates.append(rate_num)
                print(f"   ✅ Valid USD rate: ${rate_num}")
        except:
            continue
    
    if valid_rates:
        highest_rate = max(valid_rates)
        bill_rates['current_budget_rate'] = f"{highest_rate:.2f}"
        print(f"🎯 SET Current Budget rate to: ${bill_rates['current_budget_rate']}")
    
    # === Determine final rate ===
    print("\n" + "=" * 60)
    print("📊 RATE DECISION:")
    
    q_rate = float(bill_rates['questions_rate']) if bill_rates['questions_rate'] != "00" else 0
    c_rate = float(bill_rates['current_budget_rate']) if bill_rates['current_budget_rate'] != "00" else 0
    
    print(f"   Questions rate: ${q_rate}")
    print(f"   Current Budget rate: ${c_rate}")
    
    # Priority: Questions > Current Budget
    if q_rate > 0:
        bill_rates['final_bill_rate'] = bill_rates['questions_rate']
        bill_rates['source'] = 'questions'
        print(f"🏆 SELECTED: Questions rate ${bill_rates['questions_rate']}")
    elif c_rate > 0:
        bill_rates['final_bill_rate'] = bill_rates['current_budget_rate']
        bill_rates['source'] = 'current_budget'
        print(f"🥈 SELECTED: Current Budget rate ${bill_rates['current_budget_rate']}")
    else:
        bill_rates['final_bill_rate'] = "00"
        bill_rates['source'] = 'none'
        print("❌ NO RATE FOUND - using $00")
    
    print("=" * 60)
    return bill_rates

def save_extracted_data_to_file(file_path, extracted_data):
    """
    Save extracted data in the CLEAN FORMAT with separate bill rates
    """
    import os
    
    # Extract requisition number from filename
    requisition_number = "unknown"
    filename = os.path.basename(file_path)
    if "requisition_" in filename and "_complete.txt" in filename:
        try:
            parts = filename.split('_')
            if len(parts) >= 2 and parts[1].isdigit():
                requisition_number = parts[1]
        except:
            pass
    
    # Build the clean output format
    output_lines = []
    
    # Header information
    output_lines.append(f"Requisition Number: {requisition_number}")
    output_lines.append(f"Work Arrangement: {extracted_data.get('work_arrangement', 'Onsite')}")
    
    # ADD TITLE HERE
    output_lines.append(f"Title/Role: {extracted_data.get('title', 'Position')}")
    
    # Worksite Address - NEW (preserve formatting if it was multi-line)
    if extracted_data.get('worksite_address'):
        worksite = extracted_data['worksite_address']
        # If the address contains " | " separator, keep it as is
        if ' | ' in worksite:
            output_lines.append(f"Worksite Address: {worksite}")
        else:
            # For multi-line addresses joined with space, keep as single line
            output_lines.append(f"Worksite Address: {worksite}")
    
    # Location
    location_parts = []
    if extracted_data.get('worksite'):
        location_parts.append(extracted_data['worksite'])
    if extracted_data.get('work_location') and extracted_data['work_location'] not in ['GL:', 'N/A']:
        location_parts.append(extracted_data['work_location'])
    
    if location_parts:
        output_lines.append(f"Location: {' | '.join(location_parts)}")
    
    # All the fields
    output_lines.append(f"No. of Openings: {extracted_data.get('openings', '1')}")
    output_lines.append(f"Max Submittals by Vendor: {extracted_data.get('max_submittals', '1')}")
    output_lines.append(f"Total No. Filled: {extracted_data.get('filled', '0')}")
    output_lines.append(f"Start Date: {extracted_data.get('start_date', '')}")
    output_lines.append(f"End Date: {extracted_data.get('end_date', '')}")
    
    # Deadline - format as MM/DD if in MMDD format
    deadline = extracted_data.get('deadline', '')
    if deadline and len(deadline) == 4 and deadline.isdigit():
        # Format from MMDD to MM/DD
        formatted_deadline = f"{deadline[:2]}/{deadline[2:]}"
    else:
        formatted_deadline = deadline
    output_lines.append(f"No New Submittals After: {formatted_deadline}")
    
    output_lines.append(f"Current Budget Rate: ${extracted_data.get('current_budget_rate', '00')}/hour")
    output_lines.append(f"Questions Section Rate: ${extracted_data.get('questions_rate', '00')}/hour")
    output_lines.append(f"Short Description Rate: ${extracted_data.get('short_desc_rate', '00')}/hour")
    output_lines.append(f"Final Bill Rate: ${extracted_data.get('final_bill_rate', '00')}/hour (from {extracted_data.get('source', 'unknown')})")
    
    # Descriptions
    if extracted_data.get('short_description'):
        output_lines.append("SHORT DESCRIPTION:")
        output_lines.append(extracted_data['short_description'])
        output_lines.append("")  # Empty line
    
    if extracted_data.get('complete_description'):
        output_lines.append("COMPLETE DESCRIPTION:")
        output_lines.append(extracted_data['complete_description'])
        output_lines.append("")  # Empty line
    
    # Skills Table
    if extracted_data.get('skills_table'):
        output_lines.append("SKILLS TABLE:")
        output_lines.append(extracted_data['skills_table'])
    
    # Save to file
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))
    
    print(f"  💾 Saved clean formatted data to: {os.path.basename(file_path)}")
    return '\n'.join(output_lines)
 
def process_single_file_regex_extraction(file_path):
    """
    Process a single file: RAW → REGEX EXTRACT → CLEAN FORMAT
    """
    try:
        # Read original scraped data
        with open(file_path, 'r', encoding='utf-8') as f:
            original_content = f.read()
        
        original_size = len(original_content)
        filename = os.path.basename(file_path)
        
        print(f"  🔍 Extracting: {filename}")
        print(f"    📊 Original: {original_size:,} chars")
        
        # Skip if already extracted (check for new format)
        if "Requisition Number:" in original_content:
            print("    ⏩ Already extracted with clean format, skipping")
            return True
        
        # Extract data with regex
        extracted_data = extract_essential_data_with_regex(original_content)
        
        # Save in CLEAN FORMAT (not JSON)
        clean_content = save_extracted_data_to_file(file_path, extracted_data)
        
        cleaned_size = len(clean_content)
        reduction = ((original_size - cleaned_size) / original_size) * 100
        
        print(f"    ✅ Reduced by: {reduction:.1f}%")
        print(f"    📁 File updated with clean format: {filename}")
        
        return True
        
    except Exception as e:
        print(f"    ❌ Error extracting {os.path.basename(file_path)}: {str(e)}")
        return False

def extract_all_scraped_files():
    """
    Extract ALL scraped files: RAW → REGEX EXTRACT → SAME FILE
    This runs immediately after scraping is complete
    """
    output_dir = "requisition_outputs"
    
    # Get all requisition files
    requisition_files = [os.path.join(output_dir, f) for f in os.listdir(output_dir) 
                        if f.endswith('.txt') and f.startswith('requisition_')]
    
    if not requisition_files:
        print("No scraped files found for extraction")
        return 0
    
    print(f"\n🔍 REGEX EXTRACTION: {len(requisition_files)} files")
    print("=" * 50)
    
    processed_count = 0
    
    for i, file_path in enumerate(requisition_files):
        success = process_single_file_regex_extraction(file_path)
        if success:
            processed_count += 1
        
        # Add small delay between files
        if i < len(requisition_files) - 1:
            time.sleep(1)
    
    print("=" * 50)
    print(f"✅ EXTRACTION COMPLETE: {processed_count}/{len(requisition_files)} files")
    
    return processed_count



# ===== UPDATED LLM WITHOUT ANY TABLE FORMATTING IN SKILLS =====

class PureLLMRequisitionProcessor:
    def __init__(self, api_key=None, model="llama-3.1-8b-instant"):
        self.api_key = api_key or os.getenv('GROQ_API_KEY')
        self.model = model
        self.base_url = "https://api.groq.com/openai/v1/chat/completions"
        self.max_retries = 3
        self.retry_delay = 2

    def make_llm_call(self, prompt, system_message, max_tokens=2000, timeout=30):
        """Make LLM call"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": system_message
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.1,
            "max_tokens": max_tokens,
            "top_p": 0.9
        }
        
        for attempt in range(self.max_retries):
            try:
                print(f"  📡 LLM Call Attempt {attempt + 1}...")
                response = requests.post(self.base_url, headers=headers, json=payload, timeout=timeout)
                
                if response.status_code == 429:
                    retry_after = int(response.headers.get('Retry-After', 30))
                    print(f"  ⏳ Rate limited. Waiting {retry_after} seconds...")
                    time.sleep(retry_after)
                    continue
                    
                response.raise_for_status()
                result = response.json()
                
                if 'choices' in result and len(result['choices']) > 0:
                    print("  ✅ LLM call successful")
                    return result['choices'][0]['message']['content']
                else:
                    print("  ❌ No choices in response")
                    return None

            except requests.exceptions.Timeout:
                print(f"  ⏰ Timeout on attempt {attempt + 1}")
                if attempt < self.max_retries - 1:
                    time.sleep(10 * (attempt + 1))
                continue
                    
            except requests.exceptions.RequestException as e:
                print(f"  🔌 API connection error: {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(5 * (attempt + 1))
                continue

        print("  ❌ All LLM attempts failed")               
        return None
    def get_deadline_date(self, clean_content):
        """
        Get the deadline date - with proper Virginia/NC logic
        """
        # First check if it's Virginia
        is_virginia = self.is_virginia_requisition(clean_content)
    
        if is_virginia:
            print("  🏛️  This is a Virginia requisition")
            # Virginia: Calculate 4 business days
            return self.calculate_4_business_days()
        else:
            print("  🌟 This is NOT Virginia (likely NC or other)")
            # Non-Virginia: Extract from "No New Submittals After:"
            return self.extract_date_mmd(clean_content)

    def is_virginia_requisition(self, clean_content):
        """Check if this is a Virginia requisition"""
        import re
    
        # Virginia indicators
        va_indicators = [
            r'VA-\d+',  # VA-12345 pattern
            r'Worksite Address:.*,\s*VA\s*\d',
            r'Location:.*,\s*VA\s*\d',
            r'Virginia',
        ]
    
        for pattern in va_indicators:
            if re.search(pattern, clean_content, re.IGNORECASE):
                return True
    
        return False

    def calculate_4_business_days(self):
        """Calculate 4 business days from today"""
        from datetime import datetime, timedelta
    
        today = datetime.now()
        business_days_added = 0
        current_date = today
    
        while business_days_added < 4:
            current_date += timedelta(days=1)
            # Check if it's a weekday (Monday=0, Sunday=6)
            if current_date.weekday() < 5:
               business_days_added += 1
    
        mmdd = current_date.strftime("%m%d")
        print(f"  📅 Calculated 4 business days: {mmdd}")
        return mmdd


    def extract_date_mmd(self, clean_content):
        """
        Manually extract and format date as MMDD from "No New Submittals After:" field
        Returns: MMDD string (e.g., "1215") or "0101" if not found
        """
        import re
        # FIRST: Check if the field exists at all
        if "No New Submittals After:" not in clean_content:
            print("  ⚠️  'No New Submittals After' NOT FOUND in text")
            return "0101"
    
        print("  ✅ 'No New Submittals After' IS PRESENT in text")
        # Try multiple patterns to match different date formats
        patterns = [
            # Pattern 1: MM/DD or MM-DD (NO YEAR) - THIS IS YOUR NC FORMAT!
            r'No New Submittals After:\s*(\d{1,2})[-/](\d{1,2})(?:\D|$)',
        
            # Pattern 2: MM/DD/YYYY or MM-DD-YYYY (with year)
            r'No New Submittals After:\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})',
        
            # Pattern 3: Just month and day separated by space
            r'No New Submittals After:\s*(\d{1,2})\s+(\d{1,2})',
        ]
        for pattern in patterns:
            date_match = re.search(pattern, clean_content)
            if date_match:
                month = date_match.group(1).zfill(2)  # Ensure 2 digits
                day = date_match.group(2).zfill(2)    # Ensure 2 digits
                mmdd = f"{month}{day}"
            
                if len(date_match.groups()) >= 3 and date_match.group(3):
                    print(f"  📅 Extracted date with year: {mmdd}")
                else:
                    print(f"  📅 Extracted date without year: {mmdd}")
            
                return mmdd
        # If we get here, we found the field but couldn't parse the date
        print("  ⚠️  'No New Submittals After' found but COULD NOT EXTRACT DATE")
        return "0101"  

    def format_extracted_data(self, clean_content, filename, title):
        """
        Format the extracted data into exact required format WITH TITLE IN CORRECT PLACE
        """
        # FIX: Call self.extract_date_mmd() not extract_deadline_date()
        date_mmd = self.get_deadline_date(clean_content)  # Use new logic
    
        bill_rates = extract_bill_rates_from_all_sections(clean_content)
    
        final_bill_rate = bill_rates.get('final_bill_rate', '00').split('.')[0]
        coded_id = f"9{final_bill_rate}9{date_mmd}"
    
        print(f"  💰 Using final bill rate: ${final_bill_rate}")
        print(f"  📅 Using date: {date_mmd}")
        print(f"  🆔 Job ID part will be: {coded_id}")
    

        prompt = f"""
EXTRACTED JOB DATA:
{clean_content}

JOB TITLE TO USE: {title}

STRICTLY FOLLOW FORMAT THIS DATA EXACTLY AS SHOWN BELOW - NO ANALYSIS, NO EXPLANATIONS, NO ADDITIONAL TEXT:

Job ID: [State]-[RequisitionNumber] ({coded_id})

Location: [City, State (Work Location)]
Duration: [X Months]
Positions:[number_of_openings (max_submittals)]

skills:
[Skills as plain text list - NO TABLE FORMATTING]

Description:
[Combined description content without SHORT/COMPLETE headings]

CRITICAL INSTRUCTIONS - MUST FOLLOW EXACTLY:

1. OUTPUT ONLY THE FORMATTED RESULT - NO INTRODUCTORY TEXT, NO ANALYSIS, NO EXPLANATIONS
2. USE THIS EXACT TITLE: "{title}" - DO NOT MODIFY IT

3. JOB ID:
   - STATE: Find the "Worksite Address:" and extract the two-letter state abbreviation.
   - Extract the FULL requisition number (all digits) don't change 
   - FORMAT MUST BE EXACTLY: [State]-[Requisition Number] ({coded_id})
   - REMOVE DECIMAL POINTS AND ANY VALUES AFTER THEM 
   - EXAMPLE: "GA-788565 (98291216)"
   - DO NOT change the coded ID - use it exactly as provided: {coded_id}
   - EXAMPLES:
        - "85.50" → "85"
        - "137.00" → "137" 
        - "21.75" → "21"
        - "81.73" → "81"

4. LOCATION:
   - Find "Location:" field for city and state
   - Extract Work Location exactly 
   - Find the "Worksite Address:" field. Extract the city and state (2-letter code). Exclude the zip code.
   - Find the "Work Location:" field or similar in the descripthon section and extract the work location (e.g., "NCDHHS-NCFAST").
   - Format: "City, State (Work Location)" 

5. DURATION:
   - Find "Start Date:" and "End Date:" 
   - Calculate months between dates
   - Format: "X Months" (e.g., "12 Months")

6. POSITIONS:
   - Find "No. of Openings:" and also Max Submittals by Vendor and use that numbers
   - Format as "number_of_openings (Max Submittals by Vendor)" 

7. SKILLS:
   - Find the "SKILLS TABLE:" section and extract from there don't change anything like format and all 

8. DESCRIPTION:
   - COMBINE both "SHORT DESCRIPTION:" and "COMPLETE DESCRIPTION:" content
   - REMOVE the headers "SHORT DESCRIPTION:" and "COMPLETE DESCRIPTION:"
   - Keep all the actual description content
   - AFTER EACH PERIOD (.), START A NEW LINE FOR THE NEXT SENTENCE without line gap
   - REMOVE any analysis text like "Based on the provided job requisition..."
   UNIVERSAL DESCRIPTION SPACING:
   - IDENTIFY ALL section headers in the description using these rules:
     * Lines ending with ":" that are NOT continuations of sentences
     * Lines that are clearly section titles (standalone, not part of paragraphs)
     * Lines containing typical header words (Responsibilities, Qualifications, Requirements, Skills, Experience, Education, Duties, Overview,Preferred technical Skills etc.)
     * Lines that are formatted as headers (ALL CAPS, bold, underlined, or visually distinct)
     * FOR EACH SECTION HEADER: Insert exactly ONE blank line BEFORE the header
     * NO blank line AFTER the header - content should start immediately on the next line8iuuuuuuioop['
     ]
     * REMOVE ALL other blank lines from the entire description
     DESCRIPTION SPACING RULES:
     - IDENTIFY ALL section headers in the description (lines ending with ":" that are section titles)
     - FOR EACH SECTION HEADER: Insert exactly ONE blank line BEFORE the header
     - REMOVE the blank line immediately AFTER each header (content must start on the very next line with NO gap)
     - REMOVE ALL other blank lines from the entire description
     - PRESERVE all bullet points and original formatting
     - ENSURE: Header -> Immediate content (no blank line in between)

   HEADER DETECTION EXAMPLES:
   ✓ "RESPONSIBILITIES:" (header - add blank line before)
   ✓ "REQUIRED SKILLS:" (header - add blank line before)  
   ✓ "QUALIFICATIONS:" (header - add blank line before)
   ✓ "Work Environment: Professional office" (NOT a header - part of sentence)
   ✓ "Position: Compliance Officer" (NOT a header - part of sentence)
   - REMOVE ALL other blank lines from the entire description, except for the single blank lines inserted before headers as specified.
   - PRESERVE all original formatting, bullet points, and header names.
   - DO NOT add headers that aren't in the original text.
   

  FINAL RESULT MUST HAVE:
   - One blank line before each section header no blank line after header
   - AFTER EACH PERIOD (.), START A NEW LINE FOR THE NEXT SENTENCE without line gap
   - No other blank lines anywhere in the description
   - All original content preserved 
      CRITICAL: REMOVE ALL VENDOR, BILL, RATE, AND COST INFORMATION:
   - "Pay Rate: $XX.XX"
   - "Vendor Rate: $XX.XX"
   - "Maximum Vendor Submittal Rate is XX.XX/hr"
   - "Max Vendor Submittal Rate is XX.XX/hr"
   - "Maximum Vendor Submittal Rate is $XX.XX/hr"
   - "Max Vendor Submittal Rate is $XX.XX/hr"
   - "The max rate for this position is $XXX/hour
   - "Vendor Submittal Rate: $XX.XX"
   - "Bill Rate: $XX.XX"
   - "Hourly Rate: $XX.XX"
   - "Rate: $XX.XX"
   - "Cost: $XX.XX"
   - "Budget Rate: $XX.XX"
   - "is $XXX/hour"
   - "is $XXX/hr"
   - "Assignment Duration:"
   - "Work Environment:"
   - Any line containing "$" followed by numbers (e.g., "$76/hr", "$155.00/hour")
   - Any line containing "/hr" or "/hour" or "per hour"
   - Any line containing "rate" and numbers (e.g., "rate is 76/hr", "rate of $155")
   - Any line containing "submittal" and "rate" (e.g., "vendor submittal rate")
   
   SPECIFIC EXAMPLES TO REMOVE:
   - "System Analyst 3 Maximum Vendor Submittal Rate is 76/hr."
   - "Maximum Vendor Submittal Rate is $155.00/hour"
   - "Vendor Rate: $68.07/hour"
   - "Pay Rate: $85.50"
   - "Bill rate: $73.79"
   
   REMOVE THE ENTIRE LINE if it contains any of these patterns, even if it's part of a longer sentence.
Output your response in exactly this format with no additional text:
Job ID: [value]

title: [value]

Location: [value]
Duration: [value]
Positions: [value]

skills:
 [value]

Description:
[value]

No line gaps between Location,Duration,Position 
"""

        system_message = """You are a strict data formatter. You output ONLY the formatted result in the exact structure requested.

RULES:
1. NEVER add introductory text, analysis, or explanations
2. NEVER modify the provided title - use it exactly as given
3. Extract data precisely from the provided content
4. For skills: keep all the actual skills
5. For description: Remove any vendor or bill rate text and keep the actual job description
6. Output must match the exact format and structure shown
7. AFTER EACH PERIOD (.), START THE NEXT SENTENCE ON A NEW LINE WITHOUT ADDING ANY BLANK LINES without line gap
  - REMOVE the blank line immediately AFTER each header (content must start on the very next line with NO gap)
  - REMOVE ALL other blank lines from the entire description
8. JOB ID FORMAT MUST BE EXACTLY: [State]-[RequisitionNumber] (9{final_bill_rate}9{date_mmd})  
9. CRITICAL: Remove ALL vendor, bill, rate, and cost information from description:
   - Any line containing "$" followed by numbers (e.g., "$76/hr", "$155.00/hour")
   - Any line containing "/hr" or "/hour" or "per hour"
   - Any line containing "rate" and numbers (e.g., "rate is 76/hr")
   - Any line containing "submittal" and "rate" (e.g., "vendor submittal rate")
   - "Maximum Vendor Submittal Rate is XX.XX/hr"
   - "Max Vendor Submittal Rate is XX.XX/hr"
   - "The max rate for this position is $XXX/hour
   - "Pay Rate: $XX.XX"
   - "Vendor Rate: $XX.XX"
   - "Bill Rate: $XX.XX"
   - "is $XXX/hour"
   - "is $XXX/hr"
   - "Assignment Duration:"
   - "Work Environment:"
   - If a sentence contains rate information, remove the ENTIRE sentence
10. skills extract same as how it extracted from the regrex function 
11. REMOVE ALL other blank lines from the entire description
"""

        result = self.make_llm_call(prompt, system_message, max_tokens=4000, timeout=60)
        
        # Post-process the result to remove any unwanted analysis text
        if result:
            # Remove any lines that contain analysis keywords
            lines = result.split('\n')
            cleaned_lines = []
            analysis_keywords = ['based on', 'analysis:', 'here is', 'following', 'provided', 'job requisition']
            
            for line in lines:
                line_lower = line.lower()
                # Skip lines that look like analysis
                if any(keyword in line_lower for keyword in analysis_keywords) and not line.strip().startswith(('Job ID:', 'title:', 'Location:', 'Duration:', 'Positions:', 'skills:', 'Description:', '•')):
                    continue
                cleaned_lines.append(line)
            
            result = '\n'.join(cleaned_lines)
            
            # Ensure the title is exactly what we provided
            if f"title: {title}" not in result:
                # Find where to insert the title (after Job ID line)
                lines = result.split('\n')
                new_lines = []
                title_inserted = False

                for i, line in enumerate(lines):
                    new_lines.append(line)
                    if line.startswith('Job ID:') and not title_inserted:
                        # Remove spaces within the Job ID
                        if ' ' in line:
                            # Remove spaces between numbers in Job ID but keep the main structure
                            line = line.replace(' (9 ', ' (9').replace(' 9', '9').replace(' )', ')')
                            new_lines[-1] = line  # Replace the last added line

                        # Insert title line after Job ID line (without "title:" label)
                        new_lines.append('')  # Empty line for spacing
                        new_lines.append(title)
                        title_inserted = True
            
                result = '\n'.join(new_lines)

            # Remove any "title:" labels that might be in the output   
            result = result.replace(f"title: {title}", title)
            result = re.sub(r'^title:\s*', '', result, flags=re.MULTILINE)

        return result      
                

    def process_extracted_data(self, file_content, filename, title):
        """Process extracted data with exact formatting"""
        try:
            print(f"  🤖 Formatting extracted data: {filename}")
            
            # Format the extracted data WITH the title included
            formatted_content = self.format_extracted_data(file_content, filename, title)
            
            if formatted_content:
                print("  ✅ Data formatting complete")
                return formatted_content
            else:
                print("  ❌ LLM formatting failed")
                return None
            
        except Exception as e:
            print(f"  ❌ Error processing data: {str(e)}")
            return None


class RequisitionTitleGenerator:
    def __init__(self):
        api_key = os.getenv('GROQ_API_KEY')
        self.client = Groq(api_key=api_key)
        self.model = "llama-3.1-8b-instant"

    def generate_title(self, requisition_content: str) -> Optional[str]:
        """Generate title - LLM finds skills from both sections naturally"""
        prompt = f"""ANALYZE THIS JOB REQUISITION AND GENERATE A TECHNICAL TITLE:

REQUISITION CONTENT:
{requisition_content}

CRITICAL INSTRUCTIONS - MUST FOLLOW EXACTLY:

        1. FORMAT: [Work Arrangement]/Local [Job Title] (Experience+ Certifications) with [Technical Skills] experience

        2. WORK ARRANGEMENT: Find the "Work Arrangement:" field in the content and use that exact value

        3. JOB TITLE: Extract clean technical title only (remove anything after "/")
           - CREATE job title by analyzing the technical skills from BOTH the Skills Table AND Description sections.
           - Analyze what the person will actually DO based on the skills mentioned
           - Create an appropriate technical job title

        4. CERTIFICATIONS (IN PARENTHESES ONLY):
           - ONLY include actual certification names
           - NEVER include experience descriptions, skills, or qualifications
           - Example certifications: Salesforce Technical Architect, AWS Certified, PMP, CISSP
           - If no certifications found, use: (Experience+)

        5. TECHNICAL SKILLS (AFTER "with" ONLY):
           - EXTRACT ONLY TECHNICAL/TECHNOLOGY SKILLS
           - ABSOLUTELY NO SOFT SKILLS: NO "critical thinking", "teamwork", "communication", "problem-solving", "organizational skills"
           - ONLY: programming languages, frameworks, tools, platforms, systems, protocols, ServiceNow, Genetec, SQL Server, MS Excel, .NET, AWS, Azure
           - Examples: Java, Python, SQL, AWS, Azure, Docker, Kubernetes, React, .NET, HL7, Windows Server,SSRS/SSAS/SSIS
           - EXTRACT ONLY programming languages, tools, platforms, systems from Skills Table
           - ABSOLUTELY NO SOFT SKILLS, NO DESCRIPTIONS, NO REQUIREMENTS
           - REMOVE phrases like: "experience in", "knowledge of", "ability to", "understanding of"
           - REMOVE job descriptions like: "IT experience in distributed environment", "application development", "security measures"
           - Include relevant technical skills ONLY
           - FORMAT: Comma-separated technical skills only

        6. EXPERIENCE YEARS:
           - If title contains "Lead", "Senior", "Manager", "Architect", or "Project Manager": (15+)
           - Otherwise: (12+)

        7. OUTPUT MUST BE EXACT FORMAT: [Work Arrangement]/Local [Job Title] (Experience+ Certifications) with [Technical Skills] experience

        SKILLS (Go after "with"):
        - Abilities, technologies, or knowledge areas
        - Examples: Python, Java, JavaScript, React, Node.js, Angular, SQL, NoSQL, Docker, Kubernetes, AWS, Azure, GCP, Splunk, CrowdStrike, Nessus, Tableau, Power BI, machine learning, data analysis, network security, vulnerability management, restorations, dentures, extractions, patient care, dental procedures, financial analysis, risk assessment, project management, agile methodology, EagleSoft Electronic Oral Health Record, Salesforce, ServiceNow, Epic Systems, SAP, Oracle,SSRS/SSAS/SSIS,
        
        EXAMPLES OF CORRECT FORMAT:
        
        Example 1 : (Security Analyst):
        Hybrid/Local Security Analyst(12+ CISSP/CISM/Security+) with Splunk, CrowdStrike Falcon, Nessus, Tenable.sc, NIST, FISMA, vulnerability management, risk assessment experience
        
        Example 2 : (Dentist):
        Onsite/Local Dentist(12+ DDS/DMD Degree/FL Dental License/DEA/CPR) with restorations, dentures, extractions, EagleSoft Electronic Oral Health Record, patient care, dental procedures experience
        
        Example 3 : (Software Developer):
        Remote/Local Senior Full Stack Developer(15+ AWS/Azure) with Java, JavaScript, React, Node.js, Angular, SQL, NoSQL, microservices, CI/CD, Agile methodology experience
        
        Example 4 : (Salesforce Developer):
        Onsite/Local Salesforce Solutions Developer(12+ Salesforce Admin I and II/Salesforce Platform Developer I and II and/or Platform App Builder/Salesforce ServiceCloud Consultant/Copado I and II/CRT) with GITHUB, Copado, Jira, Shield, Lightning (LWC), Salesforce Flows, Salesforce Platform Community, Service Cloud, Gov Cloud, Ownbackup, Microsoft Office Suite (MS Word, EXCEL, PowerPoint, Visio), Agile methodology, Scrum, Kanban experience
        

        BAD EXAMPLES TO AVOID:
        - ❌ (12+ Hands-on experience with mobile device deployment) → REMOVE "Hands-on experience"
        - ❌ with critical thinking, teamwork, communication skills → REMOVE SOFT SKILLS
        - ❌ (Familiarity with Mobile Device Management) → NOT A CERTIFICATION

        GOOD EXAMPLES:
        - ✅ Onsite/Local Rhapsody Developer (12+) with Rhapsody Integration Engine, HL7/FHIR, .NET/C#, SQL, IIS, Windows Server, PowerShell, Oracle PL/SQL, NBS/NEDSS experience
        - ✅ Remote/Local Mobile Device Management Specialist (12+ Apple Certified Support Professional) with iOS, MDM platforms, mobile device deployment, IT security practices experience
        - ✅Remote/Local SQL Server DBA (Azure certification must) with HA/Clustering/DR, Hyper-V/VMware, T-SQL, PowerShell, Windows/Active Directory/CNO, SSRS, SSIS, SSAS, Redgate/SolarWinds/Sentry, DNS experience
        EXAMPLES OF CORRECT OUTPUT:
        - Remote/Local Service Support Analyst (12+) with ServiceNow, Genetec, SQL Server experience
        - Onsite/Local .NET Developer (12+) with .NET, SQL Server, REST APIs experience
        - Hybrid/Local Cloud Engineer (15+) with AWS, Azure, Terraform experience


        NOW GENERATE THE TITLE FOLLOWING THESE STRICT RULES:"""

        try:
            response = self.client.chat.completions.create(
                messages=[
                    {
                        "role": "system", 
                        "content": """STRICT RULES: 
                        1. OUTPUT EXACT FORMAT: [Work Arrangement]/Local [Job Title] (Experience+ Certifications) with [Technical Skills] experience
                        2. CREATE NEW JOB TITLE by analyzing skills from both Skills Table AND Description
                        4. FIND Work Arrangement from the content
                        5. CERTIFICATIONS: Only include if explicitly mentioned, otherwise use (12+) or (15+)
                        6. AFTER "with": ONLY technical skills from analysis
                        7. NO SOFT SKILLS EVER
                        8. OUTPUT ONLY THE FINAL TITLE - NO ANALYSIS, NO EXPLANATION, NO REASONING"""
                    },
                    {
                        "role": "user", 
                        "content": prompt
                    }
                ],
                model=self.model,
                temperature=0.1,
                max_tokens=150
            )
            
            title = response.choices[0].message.content.strip()
            
            # STRICT CLEANING - Remove ANY analysis text
            lines = title.split('\n')
            final_title = None
            
            # Look for the line that matches our exact title format
            for line in lines:
                line = line.strip()
                # Check if this line matches our title format pattern
                if (('/Local ' in line or '/Local' in line) and 
                    ('(12+)' in line or '(15+)' in line) and 
                    'with' in line and 
                    'experience' in line):
                    final_title = line
                    break
            
            # If no exact format found, look for any line that starts with work arrangement
            if not final_title:
                work_arrangements = ['Onsite/', 'Remote/', 'Hybrid/']
                for line in lines:
                    line = line.strip()
                    if any(line.startswith(arrangement) for arrangement in work_arrangements):
                        final_title = line
                        break
            
            # If still no title found, use the first line and clean it aggressively
            if not final_title:
                first_line = lines[0].strip()
                # Remove any markdown formatting, asterisks, etc.
                first_line = re.sub(r'\*\*|\*|\[|\]|\(|\)|\#', '', first_line)
                # Remove common analysis phrases
                analysis_phrases = [
                    'Based on the provided job requisition',
                    'here is the analysis',
                    'Work Arrangement:',
                    'Job Title:',
                    'Explanation:',
                    'Final Output:',
                    'Technical Skills:',
                    'Certifications:'
                ]
                for phrase in analysis_phrases:
                    first_line = first_line.replace(phrase, '').strip()
                final_title = first_line
            
            # Final cleanup
            if final_title:
                final_title = self.clean_generated_title(final_title)
                print(f"  🎯 Generated Title: {final_title}")
                return final_title
            else:
                print("  ❌ Title generation failed")
                return None
                
        except Exception as e:
            print(f"Generation error: {e}")
            return None

    def generate_title_from_extracted_data(self, extracted_content: str) -> Optional[str]:
        """Generate title from extracted data - wrapper around generate_title"""
        return self.generate_title(extracted_content)

    def clean_generated_title(self, title: str) -> str:
        """Clean up the generated title to enforce rules"""
        # Remove any markdown formatting
        title = re.sub(r'\*\*|\*|__|_|\[|\]|#', '', title)
        
        # Remove common analysis phrases
        analysis_phrases = [
            'Based on the provided job requisition',
            'here is the analysis',
            'Work Arrangement:',
            'Job Title:',
            'Explanation:',
            'Final Output:',
            'Technical Skills:',
            'Certifications:',
            'has the analysis',
            'the analysis is',
            'following title:',
            'generated title:',
            'title is:'
        ]
        
        for phrase in analysis_phrases:
            title = title.replace(phrase, '').strip()
        
        # Remove anything in parentheses that's not certifications or experience
        if '(' in title and ')' in title:
            # Keep only (12+), (15+), or actual certifications
            paren_content = title.split('(')[1].split(')')[0]
            certification_indicators = ['certified', 'certification', 'license', 'credential', 'pmp', 'cissp', 'aws', 'azure', 'salesforce']
            if not any(indicator in paren_content.lower() for indicator in certification_indicators) and not any(exp in paren_content for exp in ['12+', '15+']):
                title = title.replace(f"({paren_content})", "").strip()
        
        # Clean the job title - remove anything after "/" in the job title part
        if "Local " in title:
            # Split the title into work arrangement and the rest
            parts = title.split("Local ", 1)
            if len(parts) == 2:
                work_arrangement = parts[0]  # "Onsite/" or "Remote/" etc.
                rest_of_title = parts[1]     # "Automated SW Tester/Analyst (12+) with..."
                
                # Clean the job title (part before parentheses)
                if '(' in rest_of_title:
                    job_title_part = rest_of_title.split('(', 1)[0].strip()
                    remaining_part = rest_of_title.split('(', 1)[1]
                    
                    # Remove anything after "/" in the job title
                    if '/' in job_title_part:
                        job_title_part = job_title_part.split('/')[0].strip()
                    
                    # Reconstruct the title
                    title = f"{work_arrangement}Local {job_title_part} ({remaining_part}"
        
        # Remove soft skills that might have slipped through
        soft_skills = [
            "critical thinking", "teamwork", "communication", "problem-solving",
            "organizational skills", "collaboration", "customer service",
            "attention to detail", "analytical skills", "time management",
            "interpersonal skills", "leadership", "mentoring", "training",
            "documentation", "presentation skills", "written communication",
            "verbal communication", "multitasking", "adaptability", "creativity",
            "strategic thinking", "decision making", "conflict resolution",
            "negotiation", "mentoring", "coaching", "facilitation"
        ]
        
        for skill in soft_skills:
            title = title.replace(f", {skill}", "").replace(f"{skill}, ", "")
        
        # Ensure certifications are only included if actually mentioned
        # If parentheses contain non-certification content, replace with just experience
        if '(' in title and ')' in title:
            paren_content = title.split('(')[1].split(')')[0]
            # Check if this looks like actual certifications or just experience
            certification_indicators = ['certified', 'certification', 'license', 'credential', 'pmp', 'cissp', 'aws', 'azure', 'salesforce']
            has_real_certifications = any(indicator in paren_content.lower() for indicator in certification_indicators)
            
            if not has_real_certifications and 'experience' in paren_content.lower():
                # Replace with just experience years
                if any(word in title.lower() for word in ['lead', 'senior', 'manager', 'architect', 'project manager']):
                    title = title.replace(f"({paren_content})", "(15+)")
                else:
                    title = title.replace(f"({paren_content})", "(12+)")
        
        # Ensure proper spacing
        title = re.sub(r'\s+', ' ', title)
        
        return title.strip()




def process_files_with_llm():
    """Process extracted data with exact formatting"""
    print("\n=== PROCESSING EXTRACTED DATA WITH EXACT FORMATTING ===")
    
    processor = PureLLMRequisitionProcessor()
    title_generator = RequisitionTitleGenerator()
    
    output_dir = "requisition_outputs"
    requisition_files = [os.path.join(output_dir, f) for f in os.listdir(output_dir) 
                        if f.endswith('.txt') and f.startswith('requisition_')]
    
    for i, file_path in enumerate(requisition_files):
        filename = os.path.basename(file_path)
        print(f"\n📄 Processing {i+1}/{len(requisition_files)}: {filename}")
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                file_content = f.read()
            
            # Check if already processed (has "Job ID:" header)
            if "Job ID:" in file_content and "title:" in file_content:
                print("  ✅ Already formatted, skipping")
                continue
            elif "Requisition Number:" in file_content:
                print("  ✅ Processing extracted data")
                
                # STEP 1: GENERATE TITLE
                print("  🎯 Generating title...")
                title = title_generator.generate_title(file_content)
                
                if title:
                    print(f"  ✅ Final Title: {title}")
                else:
                    print("  ❌ Title generation failed")
                    title = "Position"
                
                # STEP 2: FORMAT CONTENT WITH ONLY THE FINAL TITLE
                print("  📝 Formatting content with title in correct place...")
                formatted_content = processor.process_extracted_data(file_content, filename, title)
                
                if formatted_content:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(formatted_content)
                    print("  ✅ Exact formatting complete")
                else:
                    print("  ⚠️ Formatting failed, using fallback")
                    # Fallback with clean title only
                    fallback_content = f"""Job ID: [Need to extract]

title: {title}

Location: [Need to extract]
Duration: [Need to extract]
Positions: [Need to extract]

skills:
[Need to extract skills table]

Description:
[Need to extract description]"""
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(fallback_content)
            else:
                print("  ⚠️ No extracted data found, skipping")
                continue
            
            if i < len(requisition_files) - 1:
                time.sleep(3)
                
        except Exception as e:
            print(f"  ❌ Error: {str(e)}")


def add_title_to_formatted_content(formatted_content, title):
    """Add the generated title to the formatted content"""
    if not title or not formatted_content:
        return formatted_content
    
    # Insert title at the beginning with proper spacing
    return f"{title}\n\n{formatted_content}"
 

# In the email sending section of main(), update the file finding logic:
def find_requisition_files(req_id, updated_docs_dir, state_abbr):
    """Find SM and RTR files for a given requisition ID - enhanced for combined documents"""
    sm_files = []
    rtr_files = []
    
    if not os.path.exists(updated_docs_dir):
        print(f"  ❌ Directory not found: {updated_docs_dir}")
        return sm_files, rtr_files
    
    print(f"  🔍 Searching for files for requisition {req_id} (State: {state_abbr})")
    
    # FIRST: Check for combined document (like RTR_SM_Idaho_776597.docx)
    combined_patterns = [
        f"RTR_SM_{state_abbr}_{req_id}.docx",
        f"SM_RTR_{state_abbr}_{req_id}.docx", 
        f"Combined_{state_abbr}_{req_id}.docx"
    ]
    
    for pattern in combined_patterns:
        combined_path = os.path.join(updated_docs_dir, pattern)
        if os.path.exists(combined_path):
            print(f"  ✅ Found combined document: {pattern}")
            # For combined documents, return the same file for both SM and RTR
            return [pattern], [pattern]
    
    # SECOND: Look for separate files with exact state matching
    for file in os.listdir(updated_docs_dir):
        if req_id in file and file.endswith('.docx'):
            # Exact match for state abbreviation
            if file.startswith(f'SM_{state_abbr}_'):
                sm_files.append(file)
                print(f"  ✅ Found SM file: {file}")
            elif file.startswith(f'RTR_{state_abbr}_'):
                rtr_files.append(file)
                print(f"  ✅ Found RTR file: {file}")
    
    # THIRD: If no exact matches, look for files with state abbreviation anywhere in filename
    if not sm_files or not rtr_files:
        for file in os.listdir(updated_docs_dir):
            if req_id in file and file.endswith('.docx') and state_abbr in file.upper():
                if 'SM' in file.upper() and file not in sm_files:
                    sm_files.append(file)
                    print(f"  ✅ Found SM file (loose match): {file}")
                elif 'RTR' in file.upper() and file not in rtr_files:
                    rtr_files.append(file)
                    print(f"  ✅ Found RTR file (loose match): {file}")
    
    # FOURTH: If still no files, look for any files with this req_id (as fallback)
    if not sm_files or not rtr_files:
        for file in os.listdir(updated_docs_dir):
            if req_id in file and file.endswith('.docx'):
                if 'SM' in file.upper() and file not in sm_files:
                    sm_files.append(file)
                    print(f"  ⚠️ Found SM file (req_id only): {file}")
                elif 'RTR' in file.upper() and file not in rtr_files:
                    rtr_files.append(file)
                    print(f"  ⚠️ Found RTR file (req_id only): {file}")
    
    print(f"  📊 Search results: {len(sm_files)} SM files, {len(rtr_files)} RTR files")
    return sm_files, rtr_files

def check_documents_folder(documents_dir):
    """
    Check what documents are available in the Documents folder
    """
    print("\n=== CHECKING DOCUMENTS FOLDER ===")
    if not os.path.exists(documents_dir):
        print(f"ERROR: Documents folder '{documents_dir}' does not exist!")
        return
    
    files = os.listdir(documents_dir)
    sm_files = [f for f in files if f.startswith('SM_') and f.endswith('.docx')]
    rtr_files = [f for f in files if f.startswith('RTR_') and f.endswith('.docx')]
    
    print("Available SM files:")
    for sm_file in sm_files:
        print(f"  - {sm_file}")
    
    print("Available RTR files:")
    for rtr_file in rtr_files:
        print(f"  - {rtr_file}")
    
    # Check if we have templates for all states
    for state_abbr, state_info in STATE_MAPPING.items():
        if state_abbr == 'DEFAULT':
            continue
            
        sm_found = any(state_abbr in f.upper() for f in sm_files)
        rtr_found = any(state_abbr in f.upper() for f in rtr_files)
        
        status = "✓" if sm_found and rtr_found else "✗"
        print(f"{status} {state_abbr}: SM={sm_found}, RTR={rtr_found}")


# ===== AUTOMATED MONITORING FUNCTION =====

class AutomatedRequisitionProcessor:
    def __init__(self, check_interval=300):  # 5 minutes
        self.check_interval = check_interval
        self.processed_email_ids = set()
    
    def start_monitoring(self):
        """Start automated monitoring"""
        print("🚀 Starting Automated Requisition Monitor")
        print("=" * 50)
        print("Monitoring for NEW 'Now Open' emails...")
        print(f"Check interval: {self.check_interval} seconds")
        print("Press Ctrl+C to stop monitoring")
        print("=" * 50)
        
        try:
            while True:
                print(f"\n🔍 Checking for NEW emails at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print("-" * 50)
                
                new_emails_found = self.check_and_process_new_emails()
                
                if new_emails_found:
                    print(f"✅ Processed new requisitions. Next check in {self.check_interval} seconds...")
                else:
                    print(f"⏳ No new emails found. Waiting {self.check_interval} seconds...")
                
                time.sleep(self.check_interval)
                
        except KeyboardInterrupt:
            print("\n🛑 Monitoring stopped by user")
        except Exception as e:
            print(f"❌ Monitoring error: {e}")
    
    def check_and_process_new_emails(self):
        """Check for and process ONLY new emails that haven't been processed before"""
        try:
            # Authenticate with Gmail
            gmail_service = auto_authenticate_primary_gmail()
            
            # Get TODAY'S emails
            today_date = datetime.now().strftime("%Y/%m/%d")
            query = f'subject:"Now Open" newer_than:1d'
            
            results = gmail_service.users().messages().list(
                userId='me',
                labelIds=['INBOX'],
                q=query
            ).execute()
            
            messages = results.get('messages', [])
            
            if not messages:
                print("📭 No 'Now Open' emails found from today")
                return False
            
            # Filter out already processed emails
            new_messages = [msg for msg in messages if msg['id'] not in self.processed_email_ids]
            
            if not new_messages:
                print("✅ No NEW emails found (all today's emails already processed)")
                return False
            
            print(f"📨 Found {len(new_messages)} NEW 'Now Open' emails from today")
            
            # Extract URLs from NEW emails only
            all_new_urls = []
            for msg in new_messages:
                msg_id = msg['id']
                print(f"📧 Processing NEW email: {msg_id}")
                
                urls = extract_urls_from_single_email(gmail_service, msg_id)
                if urls:
                    all_new_urls.extend(urls)
                    print(f"   Found {len(urls)} requisition URLs")
                
                # Mark as processed
                self.processed_email_ids.add(msg_id)
            
            if not all_new_urls:
                print("❌ No Vector VMS links found in new emails")
                return False
            
            print(f"🔄 Found {len(all_new_urls)} NEW requisition URLs total, starting processing...")
            
            # Process the new requisitions
            return self.process_requisitions(all_new_urls)
            
        except Exception as e:
            print(f"❌ Error checking new emails: {e}")
            return False
    
    def process_requisitions(self, urls):
        """Process a list of requisition URLs"""
        if not urls:
            return False
        
        print(f"🎯 PROCESSING {len(urls)} NEW REQUISITIONS")
        
        # Initialize driver
        driver = initialize_driver()
        
        try:
            # Clear previous outputs for fresh processing
            self.clear_processing_directories()
            
            # Use existing functions for processing
            open_all_requisitions_in_new_tabs(driver, urls)
            
            # Immediate regex extraction
            print("\n" + "="*60)
            print("IMMEDIATE REGEX EXTRACTION")
            print("="*60)
            extraction_count = extract_all_scraped_files()
            print(f"✅ Raw data extracted and saved: {extraction_count} files")
            
            # Process the extracted data to create documents
            print("\n=== Creating Documents ===")
            process_requisition_files()
            
            # LLM processing
            print("\n" + "="*60)
            print("ENHANCING WITH LLM")
            print("="*60)
            try:
                process_files_with_llm() 
            except Exception as llm_error:
                print(f"⚠️ LLM processing skipped: {llm_error}")
            
            # Email sending
            print("\n" + "="*60)
            print("SENDING EMAILS")
            print("="*60)
            
            # Create a wrapper function for email sending
            email_count = self.send_emails_for_current_cycle()
            
            print(f"\n✅ NEW requisitions processed COMPLETELY! Sent {email_count} emails.")
            return True
            
        except Exception as e:
            print(f"❌ Error processing requisitions: {e}")
            return False
        finally:
            try:
                driver.quit()
            except:
                pass
    
    def clear_processing_directories(self):
        """Clear processing directories for fresh start"""
        output_dir = "requisition_outputs"
        updated_docs_dir = "updated_documents"
        
        # Clear output directory
        if os.path.exists(output_dir):
            for file in os.listdir(output_dir):
                file_path = os.path.join(output_dir, file)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Error deleting file {file_path}: {e}")
            print(f"🧹 Cleared output directory: {output_dir}")
        
        # Clear updated documents directory
        if os.path.exists(updated_docs_dir):
            for file in os.listdir(updated_docs_dir):
                file_path = os.path.join(updated_docs_dir, file)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Error deleting file {file_path}: {e}")
            print(f"🧹 Cleared documents directory: {updated_docs_dir}")
    
    def send_emails_for_current_cycle(self):
        """Send emails for current processing cycle"""
        print("📧 Sending emails for current batch...")
        
        # Check email configuration first
        to_emails = validate_email_addresses(EMAIL_CONFIG.get('to_emails', []))
        cc_emails = validate_email_addresses(EMAIL_CONFIG.get('cc_emails', []))
        bcc_emails = validate_email_addresses(EMAIL_CONFIG.get('bcc_emails', []))
        
        if not any([to_emails, cc_emails, bcc_emails]):
            print("❌ EMAIL CONFIGURATION NEEDED:")
            print("   Please set up email recipients in EMAIL_CONFIG")
            return 0
        
        # Use existing email sending logic
        output_dir = "requisition_outputs"
        updated_docs_dir = "updated_documents"
        email_count = 0
        
        for filename in os.listdir(output_dir):
            if filename.startswith("requisition_") and filename.endswith("_complete.txt"):
                match = re.search(r'requisition_(\d+)_complete\.txt', filename)
                if not match:
                    continue
                    
                req_id = match.group(1)
                req_file = os.path.join(output_dir, filename)
                
                try:
                    with open(req_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    state_abbr = extract_state_from_job_id(content) or 'DEFAULT'
                    title = "Position"
                    
                    if "Job ID:" in content:
                        lines = content.split('\n')
                        for i, line in enumerate(lines):
                            if line.startswith('Job ID:'):
                                for j in range(i+1, min(i+10, len(lines))):
                                    if lines[j].strip() and not lines[j].startswith(('Location:', 'Duration:', 'Position:')):
                                        title = lines[j].strip()
                                        break
                                break
                    
                    # Find documents
                    sm_files, rtr_files = find_requisition_files(req_id, updated_docs_dir, state_abbr)
                    
                    if sm_files and rtr_files:
                        sm_file = os.path.join(updated_docs_dir, sm_files[0])
                        rtr_file = os.path.join(updated_docs_dir, rtr_files[0])
                        
                        success = send_requisition_email(
                            req_id, title, state_abbr, content, sm_file, rtr_file,
                            to_emails=to_emails, cc_emails=cc_emails, bcc_emails=bcc_emails
                        )
                        
                        if success:
                            email_count += 1
                            print(f"✅ Email sent for requisition {req_id}")
                        else:
                            print(f"❌ Failed to send email for requisition {req_id}")
                    else:
                        print(f"❌ Documents not found for requisition {req_id}")
                        
                except Exception as e:
                    print(f"❌ Error sending email for {req_id}: {e}")
        
        print(f"📤 Email sending complete: {email_count} emails sent")
        return email_count


# Keep this function for backward compatibility
def start_automated_monitoring():
    """Legacy function - now uses AutomatedRequisitionProcessor"""
    processor = AutomatedRequisitionProcessor()
    processor.start_monitoring()

def check_and_process_new_emails(processed_email_ids):
    """Check for and process ONLY new emails that haven't been processed before"""
    try:
        # Authenticate with Gmail
        gmail_service = auto_authenticate_primary_gmail()
        
        # Get TODAY'S emails
        today_date = datetime.now().strftime("%Y/%m/%d")
        query = f'subject:"Now Open" newer_than:1d'
        
        results = gmail_service.users().messages().list(
            userId='me',
            labelIds=['INBOX'],
            q=query
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            print("📭 No 'Now Open' emails found from today")
            return False
        
        # Filter out already processed emails
        new_messages = [msg for msg in messages if msg['id'] not in processed_email_ids]
        
        if not new_messages:
            print("✅ No NEW emails found (all today's emails already processed)")
            return False
        
        print(f"📨 Found {len(new_messages)} NEW 'Now Open' emails from today")
        
        # Extract URLs from NEW emails only
        all_new_urls = []
        for msg in new_messages:
            msg_id = msg['id']
            print(f"📧 Processing NEW email: {msg_id}")
            
            urls = extract_urls_from_single_email(gmail_service, msg_id)
            if urls:
                all_new_urls.extend(urls)
                print(f"   Found {len(urls)} requisition URLs")
            
            # Mark as processed
            processed_email_ids.add(msg_id)
        
        if not all_new_urls:
            print("❌ No Vector VMS links found in new emails")
            return False
        
        print(f"🔄 Found {len(all_new_urls)} NEW requisition URLs total, starting processing...")
        
        # Clear previous outputs for fresh processing
        clear_processing_directories()
        
        # Process the new requisitions
        return process_requisitions(all_new_urls)
        
    except Exception as e:
        print(f"❌ Error checking new emails: {e}")
        return False

def extract_urls_from_single_email(gmail_service, message_id):
    """Extract URLs from a single email"""
    try:
        msg_data = gmail_service.users().messages().get(
            userId='me',
            id=message_id,
            format='full'
        ).execute()
        
        payload = msg_data['payload']
        html_body = ""
        plain_text = ""
        
        # Extract email content (same logic as your existing function)
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/html':
                    data = part['body'].get('data', '')
                    if data:
                        html_body = base64.urlsafe_b64decode(data).decode('utf-8')
                elif part['mimeType'] == 'text/plain':
                    data = part['body'].get('data', '')
                    if data:
                        plain_text = base64.urlsafe_b64decode(data).decode('utf-8')
        else:
            if payload['mimeType'] == 'text/html':
                data = payload['body'].get('data', '')
                if data:
                    html_body = base64.urlsafe_b64decode(data).decode('utf-8')
            elif payload['mimeType'] == 'text/plain':
                data = payload['body'].get('data', '')
                if data:
                    plain_text = base64.urlsafe_b64decode(data).decode('utf-8')
        
        links = extract_direct_links(html_body, plain_text)
        return links
        
    except Exception as e:
        print(f"❌ Error extracting URLs from email: {e}")
        return []

def clear_processing_directories():
    """Clear processing directories for fresh start"""
    output_dir = "requisition_outputs"
    updated_docs_dir = "updated_documents"
    
    # Clear output directory
    if os.path.exists(output_dir):
        for file in os.listdir(output_dir):
            file_path = os.path.join(output_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
        print(f"🧹 Cleared output directory: {output_dir}")
    
    # Clear updated documents directory
    if os.path.exists(updated_docs_dir):
        for file in os.listdir(updated_docs_dir):
            file_path = os.path.join(updated_docs_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
        print(f"🧹 Cleared documents directory: {updated_docs_dir}")

def process_requisitions(urls):
    """Process a list of requisition URLs"""
    if not urls:
        return False
    
    print(f"🎯 PROCESSING {len(urls)} NEW REQUISITIONS")
    
    # Initialize driver
    driver = initialize_driver()
    
    try:
        # Use existing functions for processing
        open_all_requisitions_in_new_tabs(driver, urls)
        process_requisition_files()
        
        # LLM processing with better error handling
        print("\n" + "="*60)
        print("ENHANCING WITH LLM")
        print("="*60)
        try:
            enhanced_count = process_files_with_llm_enhanced()
            print(f"✅ LLM enhanced {enhanced_count} files")
        except Exception as llm_error:
            print(f"⚠️ LLM processing skipped: {llm_error}")
        
        # Email sending
        print("\n" + "="*60)
        print("SENDING EMAILS")
        print("="*60)
        send_emails_for_current_cycle()
        
        print("\n✅ NEW requisitions processed COMPLETELY!")
        return True
        
    except Exception as e:
        print(f"❌ Error processing requisitions: {e}")
        return False
    finally:
        try:
            driver.quit()
        except:
            pass

def process_files_with_llm_enhanced():
    """Enhanced LLM processing with better rate limit handling"""
    print("🤖 Starting LLM enhancement with rate limit protection...")
    output_dir = "requisition_outputs"
    enhanced_count = 0
    
    if not os.path.exists(output_dir):
        return 0
    
    requisition_files = []
    for file_name in os.listdir(output_dir):
        if file_name.endswith('.txt') and file_name.startswith('requisition_'):
            requisition_files.append(os.path.join(output_dir, file_name))
    
    if not requisition_files:
        return 0
    
    print(f"📝 Processing {len(requisition_files)} files with LLM...")
    
    for file_path in requisition_files:
        file_name = os.path.basename(file_path)
        print(f"  Processing {file_name}...")
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                original_content = f.read()
            
            # Add longer delay between LLM calls to avoid rate limits
            time.sleep(5)  # Increased from 2 to 5 seconds
            
            # Use your existing LLM processing
            processor = PureLLMRequisitionProcessor(
                api_key=os.getenv('GROQ_API_KEY'),
                model="llama-3.1-8b-instant"
            )
            title_generator = RequisitionTitleGenerator()
            
            # Generate title
            title = title_generator.generate_title(original_content)
            
            # Format content
            formatted_content = processor.process_file_with_pure_llm(file_name, original_content)
            
            if formatted_content and "failed" not in formatted_content.lower() and title:
                final_content = add_title_to_formatted_content(formatted_content, title)
                
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(final_content)
                
                enhanced_count += 1
                print(f"    ✅ Successfully enhanced {file_name}")
            else:
                print(f"    ⚠️ LLM processing limited for {file_name} - keeping original content")
                
        except Exception as e:
            print(f"    ❌ Error processing {file_name}: {str(e)}")
            # Continue with other files even if one fails
    
    return enhanced_count

def send_emails_for_current_cycle():
    """Send emails for current processing cycle"""
    print("📧 Sending emails for current batch...")
    
    # Check email configuration first
    to_emails = validate_email_addresses(EMAIL_CONFIG.get('to_emails', []))
    cc_emails = validate_email_addresses(EMAIL_CONFIG.get('cc_emails', []))
    bcc_emails = validate_email_addresses(EMAIL_CONFIG.get('bcc_emails', []))
    
    if not any([to_emails, cc_emails, bcc_emails]):
        print("❌ EMAIL CONFIGURATION NEEDED:")
        print("   Please set up email recipients in EMAIL_CONFIG")
        print("   Current config - TO:", EMAIL_CONFIG.get('to_emails'))
        print("   Current config - CC:", EMAIL_CONFIG.get('cc_emails')) 
        print("   Current config - BCC:", EMAIL_CONFIG.get('bcc_emails'))
        return 0
    
    # Use your existing email sending logic from main()
    output_dir = "requisition_outputs"
    updated_docs_dir = "updated_documents"
    email_count = 0
    
    for filename in os.listdir(output_dir):
        if filename.startswith("requisition_") and filename.endswith("_complete.txt"):
            match = re.search(r'requisition_(\d+)_complete\.txt', filename)
            if not match:
                continue
                
            req_id = match.group(1)
            req_file = os.path.join(output_dir, filename)
            
            try:
                with open(req_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                state_abbr = extract_state_from_job_id(content) or 'DEFAULT'
                title = "Position"
                
                if "Job ID:" in content:
                    lines = content.split('\n')
                    for i, line in enumerate(lines):
                        if line.startswith('Job ID:'):
                            for j in range(i+1, min(i+10, len(lines))):
                                if lines[j].strip() and not lines[j].startswith(('Location:', 'Duration:', 'Position:')):
                                    title = lines[j].strip()
                                    break
                            break
                
                # Find documents
                sm_files, rtr_files = find_requisition_files(req_id, updated_docs_dir, state_abbr)
                
                if sm_files and rtr_files:
                    sm_file = os.path.join(updated_docs_dir, sm_files[0])
                    rtr_file = os.path.join(updated_docs_dir, rtr_files[0])
                    
                    success = send_requisition_email(
                        req_id, title, state_abbr, content, sm_file, rtr_file,
                        to_emails=to_emails, cc_emails=cc_emails, bcc_emails=bcc_emails
                    )
                    
                    if success:
                        email_count += 1
                        print(f"✅ Email sent for requisition {req_id}")
                    else:
                        print(f"❌ Failed to send email for requisition {req_id}")
                else:
                    print(f"❌ Documents not found for requisition {req_id}")
                    
            except Exception as e:
                print(f"❌ Error sending email for {req_id}: {e}")
    
    print(f"📤 Email sending complete: {email_count} emails sent")
    return email_count

# Integrating the due list and its related functions 

def generate_requisition_report():
    """Simple wrapper to run the updated reporting system"""
    try:
        # Import the main function from suresh1.py instead of suresh.py
        from suresh1 import main as report_main
        report_main()
        return True
    except Exception as e:
        print(f"Report generation failed: {e}")
        return False

def extract_urls_from_single_email(gmail_service, message_id):
    """Extract URLs from a single email"""
    try:
        msg_data = gmail_service.users().messages().get(
            userId='me',
            id=message_id,
            format='full'
        ).execute()
        
        payload = msg_data['payload']
        html_body = ""
        plain_text = ""
        
        # Extract email content (same logic as your existing function)
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/html':
                    data = part['body'].get('data', '')
                    if data:
                        html_body = base64.urlsafe_b64decode(data).decode('utf-8')
                elif part['mimeType'] == 'text/plain':
                    data = part['body'].get('data', '')
                    if data:
                        plain_text = base64.urlsafe_b64decode(data).decode('utf-8')
        else:
            if payload['mimeType'] == 'text/html':
                data = payload['body'].get('data', '')
                if data:
                    html_body = base64.urlsafe_b64decode(data).decode('utf-8')
            elif payload['mimeType'] == 'text/plain':
                data = payload['body'].get('data', '')
                if data:
                    plain_text = base64.urlsafe_b64decode(data).decode('utf-8')
        
        links = extract_direct_links(html_body, plain_text)
        return links
        
    except Exception as e:
        print(f"❌ Error extracting URLs from email: {e}")
        return []

def send_excel_report_email(excel_file_path, subject_prefix="Requisition Due List Report"):
    """Send the Excel report as a separate email to multiple recipients"""
    
    try:
        # Validate email addresses
        to_emails = validate_email_addresses(EMAIL_CONFIG.get('to_emails', []))
        cc_emails = validate_email_addresses(EMAIL_CONFIG.get('cc_emails', []))
        bcc_emails = validate_email_addresses(EMAIL_CONFIG.get('bcc_emails', []))
        
        if not any([to_emails, cc_emails, bcc_emails]):
            logging.error("No valid email recipients found for Excel report")
            return False
        
        sender_email = EMAIL_CONFIG['sender_email']
        password = EMAIL_CONFIG['password']
        smtp_server = EMAIL_CONFIG['smtp_server']
        smtp_port = EMAIL_CONFIG['smtp_port']
        
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = sender_email
        
        # Set recipients
        if to_emails:
            msg['To'] = ', '.join(to_emails)
        
        # Add CC if specified
        if cc_emails:
            msg['Cc'] = ', '.join(cc_emails)
        
        # Current date for subject
        current_date = datetime.now().strftime("%Y-%m-%d")
        msg['Subject'] = f"{subject_prefix} - {current_date}"
        
        # Email body
        body = f"""
        Hello Team,
        
        Please find attached the Combined Requisition Due List Report for {current_date}.
        
        This report contains:
        - All processed requisitions from Vector VMS
        - All solicitation responses from HHSC Portal
        - Due date categorization (Active vs Past Due)
        - Job titles and IDs from both sources
        
        The report is automatically generated and includes all requisitions processed today.
        
        Report Features:
        • Active Due: Jobs with future due dates
        • Past Due: Jobs with due dates that have passed
        • Combined data from multiple sources
        
        Best regards,
        Automated Requisition System
        """
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach Excel file - UPDATED FILENAME
        excel_filename = "combined_due_list.xlsx"  # Updated from requisition_due_list_report.xlsx
        if os.path.exists(excel_filename):
            try:
                with open(excel_filename, 'rb') as file:
                    part = MIMEApplication(file.read(), Name=excel_filename)
                
                part['Content-Disposition'] = f'attachment; filename="{excel_filename}"'
                msg.attach(part)
                print(f"  ✅ Excel file attached: {excel_filename}")
            except Exception as e:
                logging.error(f"Failed to attach Excel file: {str(e)}")
                return False
        else:
            # Try the old filename as fallback
            if os.path.exists(excel_file_path):
                try:
                    with open(excel_file_path, 'rb') as file:
                        part = MIMEApplication(file.read(), Name=os.path.basename(excel_file_path))
                    
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_file_path)}"'
                    msg.attach(part)
                    print(f"  ✅ Excel file attached (fallback): {os.path.basename(excel_file_path)}")
                except Exception as e:
                    logging.error(f"Failed to attach Excel file: {str(e)}")
                    return False
            else:
                logging.error(f"Excel file not found: {excel_filename} or {excel_file_path}")
                return False
        
        # Prepare all recipients (TO + CC + BCC)
        all_recipients = []
        if to_emails:
            all_recipients.extend(to_emails)
        if cc_emails:
            all_recipients.extend(cc_emails)
        if bcc_emails:
            all_recipients.extend(bcc_emails)
        
        # Remove duplicates
        all_recipients = list(set(all_recipients))
        
        # Send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, all_recipients, msg.as_string())
        
        # Log recipients
        recipient_info = f"TO: {to_emails}" if to_emails else ""
        if cc_emails:
            recipient_info += f", CC: {cc_emails}"
        if bcc_emails:
            recipient_info += f", BCC: {bcc_emails}"
        
        logging.info(f"Excel report email sent successfully to {recipient_info}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send Excel report email: {str(e)}")
        return False

def main(automated=False):
    """Main function that can run in manual or automated mode"""
    
    # Create output directory if it doesn't exist
    output_dir = "requisition_outputs"
    updated_docs_dir = "updated_documents"

    check_documents_folder("Documents")

    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
    else:
        # Clear output directory if it exists (remove all files)
        for file in os.listdir(output_dir):
            file_path = os.path.join(output_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
        print(f"Cleared output directory: {output_dir}")
    
    # Ensure updated_docs_dir exists - but don't clear it if it exists
    if not os.path.exists(updated_docs_dir):
        os.makedirs(updated_docs_dir)
        print(f"Created updated documents directory: {updated_docs_dir}")
    else:
        print(f"Using existing updated documents directory: {updated_docs_dir}")
        # Don't clear it - we want to keep all documents
    
    if automated:
        # Start automated monitoring
        processor = AutomatedRequisitionProcessor()
        processor.start_monitoring()
    else:
        # Run manual single execution (original behavior)
        # Initialize driver with enhanced options
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        try:
            # First, authenticate with Gmail and extract TODAY'S requisition URLs
            print("=== Starting Gmail Authentication ===")
            print(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
            gmail_service = auto_authenticate_primary_gmail()
            
            print("=== Extracting TODAY'S Requisition URLs ===")
            requisition_urls = extract_requisition_urls_from_gmail(gmail_service)
            
            if not requisition_urls:
                print("❌ No TODAY'S requisition URLs found in Gmail. Exiting.")
                print("✅ No new requisitions for today. Process complete.")
                return  # Exit early if no today's emails
            
            print(f"Found {len(requisition_urls)} requisition URLs from TODAY:")
            for url in requisition_urls:
                print(f"  - {url}")
            
            # Open all TODAY'S requisitions in new tabs with credential fallback logic
            open_all_requisitions_in_new_tabs(driver, requisition_urls)
            
            # ===== IMMEDIATE REGEX EXTRACTION =====
            print("\n" + "="*60)
            print("IMMEDIATE REGEX EXTRACTION")
            print("="*60)

            extraction_count = extract_all_scraped_files()
            print(f"✅ Raw data extracted and saved: {extraction_count} files")

            # Process the extracted data to create documents (uses extracted data)
            print("\n=== Creating Today's Documents ===")
            process_requisition_files()

            # LLM Enhancement Process (uses extracted data)
            print("\n" + "="*60)
            print("ENHANCING WITH LLM")
            print("="*60)

            try:
                process_files_with_llm() 
            except Exception as llm_error:
                print(f"LLM processing skipped or failed: {llm_error}")
                
            
            # ===== EMAIL SENDING SECTION =====
            print("\n" + "="*60)
            print("SENDING EMAILS")
            print("="*60)
            
            # Send emails for all processed requisitions
            email_count = 0
            
            for filename in os.listdir(output_dir):
                if filename.startswith("requisition_") and filename.endswith("_complete.txt"):
                    # Extract requisition ID
                    match = re.search(r'requisition_(\d+)_complete\.txt', filename)
                    if not match:
                        continue
                        
                    req_id = match.group(1)
                    
                    # Read the file content to extract state and title
                    req_file = os.path.join(output_dir, filename)
                    with open(req_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Extract state abbreviation from content
                    state_abbr = extract_state_from_job_id(content)
                    if not state_abbr:
                        state_abbr = 'DEFAULT'
                    
                    # Extract title from content
                    title = "Position"
                    if "Job ID:" in content:
                        lines = content.split('\n')
                        for i, line in enumerate(lines):
                            if line.startswith('Job ID:'):
                                # Look for title in subsequent lines
                                for j in range(i+1, min(i+10, len(lines))):
                                    if lines[j].strip() and not lines[j].startswith(('Location:', 'Duration:', 'Position:')):
                                        title = lines[j].strip()
                                        break
                                break
                    
                    # Build state-specific file paths (WITH STATE CODES)
                    sm_file = os.path.join(updated_docs_dir, f"SM_{state_abbr}_{req_id}.docx")
                    rtr_file = os.path.join(updated_docs_dir, f"RTR_{state_abbr}_{req_id}.docx")
                    
                    # Also check for files with "ID" in case the document functions didn't save correctly
                    sm_file_id = os.path.join(updated_docs_dir, f"SM_ID_{req_id}.docx")
                    rtr_file_id = os.path.join(updated_docs_dir, f"RTR_ID_{req_id}.docx")
                    
                    # Use state-specific files if they exist, otherwise fall back to ID files
                    files_found = True
                    if not os.path.exists(sm_file) and os.path.exists(sm_file_id):
                        sm_file = sm_file_id
                        print(f"Using SM file with ID: {os.path.basename(sm_file)}")
                    elif not os.path.exists(sm_file):
                        files_found = False
                        print(f"SM file not found for requisition {req_id}")
                    
                    if not os.path.exists(rtr_file) and os.path.exists(rtr_file_id):
                        rtr_file = rtr_file_id
                        print(f"Using RTR file with ID: {os.path.basename(rtr_file)}")
                    elif not os.path.exists(rtr_file):
                        files_found = False
                        print(f"RTR file not found for requisition {req_id}")
                    
                    # If files still don't exist, try to find any files with this req_id
                    if not files_found:
                        print(f"Searching for any files containing requisition ID {req_id}...")
                        sm_files, rtr_files = find_requisition_files(req_id, updated_docs_dir, state_abbr)
                        
                        if sm_files:
                            sm_file = os.path.join(updated_docs_dir, sm_files[0])
                            print(f"Using found SM file: {sm_files[0]}")
                            files_found = True
                        else:
                            print(f"No SM files found for requisition {req_id}")
                            files_found = False
                            
                        if rtr_files:
                            rtr_file = os.path.join(updated_docs_dir, rtr_files[0])
                            print(f"Using found RTR file: {rtr_files[0]}")
                            files_found = True
                        else:
                            print(f"No RTR files found for requisition {req_id}")
                            files_found = False
                    
                    # Check if files actually exist before sending email
                    if files_found and os.path.exists(sm_file) and os.path.exists(rtr_file):
                        # Send email with the found files
                        success = send_requisition_email(
                            req_id, 
                            title, 
                            state_abbr,
                            content, 
                            sm_file, 
                            rtr_file,
                            to_emails=EMAIL_CONFIG.get('to_emails'),
                            cc_emails=EMAIL_CONFIG.get('cc_emails'),
                            bcc_emails=EMAIL_CONFIG.get('bcc_emails')
                        )
                        
                        if success:
                            print(f"✅ Email sent for requisition {req_id} (State: {state_abbr})")
                            email_count += 1
                        else:
                            print(f"❌ Failed to send email for requisition {req_id} (State: {state_abbr})")
                    else:
                        print(f"❌ Files not found for requisition {req_id}:")
                        print(f"   SM file exists: {os.path.exists(sm_file)} - {sm_file}")
                        print(f"   RTR file exists: {os.path.exists(rtr_file)} - {rtr_file}")
                        
                        # Try one more time with any available files
                        all_sm_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('SM_') and req_id in f]
                        all_rtr_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('RTR_') and req_id in f]
                        
                        if all_sm_files and all_rtr_files:
                            sm_file = os.path.join(updated_docs_dir, all_sm_files[0])
                            rtr_file = os.path.join(updated_docs_dir, all_rtr_files[0])
                            
                            print(f"Trying with available files: {all_sm_files[0]} and {all_rtr_files[0]}")
                            
                            success = send_requisition_email(
                                req_id, 
                                title, 
                                state_abbr,
                                content, 
                                sm_file, 
                                rtr_file,
                                to_emails=EMAIL_CONFIG.get('to_emails'),
                                cc_emails=EMAIL_CONFIG.get('cc_emails'),
                                bcc_emails=EMAIL_CONFIG.get('bcc_emails')
                            )
                            
                            if success:
                                print(f"✅ Email sent for requisition {req_id} using available files")
                                email_count += 1
            
            '''# ===== UPDATED REPORTING SYSTEM =====
            print("\n" + "="*60)
            print("📊 GENERATING COMBINED REQUISITION REPORT")
            print("="*60)

            # UPDATED: Use the new combined report filename
            excel_report_path = "combined_due_list.xlsx"
            report_generated = False

            try:
                # UPDATED: Import and run the NEW reporting system from suresh1.py
                from suresh1 import main as report_main
                report_main()
                report_generated = True
                print("✅ Combined requisition report completed successfully!")
                
                # Send Excel report via email - check for the correct file
                if os.path.exists(excel_report_path):
                    print("\n" + "="*60)
                    print("📧 SENDING COMBINED EXCEL REPORT VIA EMAIL")
                    print("="*60)
                    
                    email_success = send_excel_report_email(excel_report_path, "Combined Requisition Due List Report")
                    if email_success:
                        print("✅ Combined Excel report email sent successfully!")
                    else:
                        print("❌ Failed to send combined Excel report email")
                else:
                    print(f"❌ Combined Excel report file not found: {excel_report_path}")
                    # Try to find the old file as fallback
                    old_report_path = "requisition_due_list_report.xlsx"
                    if os.path.exists(old_report_path):
                        print(f"⚠️  Found old report file, using as fallback: {old_report_path}")
                        email_success = send_excel_report_email(old_report_path, "Requisition Due List Report")
                        if email_success:
                            print("✅ Fallback Excel report email sent successfully!")
                        else:
                            print("❌ Failed to send fallback Excel report email")
                    
            except Exception as report_error:
                print(f"⚠️  Combined requisition report skipped: {report_error}")'''

            # ===== ORIGINAL FINAL SUMMARY =====
            print("\n=== Today's Processing Complete ===")
            print(f"Processed {len(requisition_urls)} requisitions from today successfully!")
            print(f"Sent {email_count} emails successfully!")

            '''if report_generated:
                if os.path.exists("combined_due_list.xlsx"):
                    print(f"📊 Combined Excel report generated and emailed: combined_due_list.xlsx")
                elif os.path.exists("requisition_due_list_report.xlsx"):
                    print(f"📊 Excel report generated and emailed: requisition_due_list_report.xlsx")
                else:
                    print("📊 Excel report was not generated")
            else:
                print("📊 Excel report was not generated")'''

            print("All today's requisitions have been processed, documents created, and emails sent!")
            
            # Show all created documents
            print("\n=== Created Documents ===")
            sm_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('SM_')]
            rtr_files = [f for f in os.listdir(updated_docs_dir) if f.startswith('RTR_')]
            
            print(f"SM Documents: {len(sm_files)}")
            for sm_file in sm_files:
                print(f"  - {sm_file}")
            
            print(f"RTR Documents: {len(rtr_files)}")
            for rtr_file in rtr_files:
                print(f"  - {rtr_file}")
        
        except Exception as main_error:
            print(f"Fatal error in processing today's requisitions: {str(main_error)}")
            import traceback
            traceback.print_exc()
            # Capture final error state
            try:
                driver.save_screenshot("final_error.png")
            except:
                pass
        
        finally:
            try:
                driver.quit()
            except:
                pass
            print("Processing complete")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--auto":
        # Automated monitoring mode
        start_automated_monitoring()
    else:
        # Single run mode (original behavior)
        print("🚀 Starting Single Requisition Processing")
        print("=" * 50)
        main()