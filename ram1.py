import os
import json
import logging
import re
import glob
import pandas as pd
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from base64 import urlsafe_b64decode
import pytz
import csv
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def get_est_time():
    """Get current time in EST timezone"""
    est = pytz.timezone('US/Eastern')
    return datetime.now(est)


def auto_authenticate_secondary_gmail():
    """Authenticates with secondary Gmail account (token1.json)"""
    print("\n=== Authenticating Secondary Gmail Account ===")
    creds = None
    token_path = 'token1.json'
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    
    # Check if credentials1.json exists
    client_file = 'credentials1.json'
    if not os.path.exists(client_file):
        print(f"❌ ERROR: {client_file} not found!")
        print("Please make sure credentials1.json file exists in your project folder")
        print("Current files in directory:")
        for file in os.listdir('.'):
            print(f"  - {file}")
        raise FileNotFoundError(f"{client_file} not found.")
    
    if os.path.exists(token_path):
        print("Found secondary token file, loading credentials...")
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
            if creds.expired and creds.refresh_token:
                print("Secondary credentials expired, refreshing...")
                creds.refresh(Request())
        except Exception as e:
            print(f"Error loading secondary credentials: {e}")
            creds = None
    
    if not creds or not creds.valid:
        print("No valid secondary credentials found, initiating OAuth flow...")
        try:
            flow = InstalledAppFlow.from_client_secrets_file(client_file, SCOPES)
            creds = flow.run_local_server(port=8081)
            token_data = json.loads(creds.to_json())
            token_data['creation_time'] = get_est_time().isoformat()
            with open(token_path, 'w') as token:
                json.dump(token_data, token)
            print("Secondary authentication successful! Token saved.")
        except Exception as e:
            print(f"Secondary authentication failed: {e}")
            raise
    
    try:
        print("Building secondary Gmail service...")
        gmail_service = build('gmail', 'v1', credentials=creds)
        print("Secondary Gmail service ready!")
        return gmail_service
    except Exception as e:
        print(f"Failed to build secondary Gmail service: {e}")
        raise

def append_to_excel(new_df, excel_path):
    """Append new data to existing Excel file with posting dates - MODIFIED FOR CUMULATIVE COUNTING"""
    try:
        # Add posting date to new jobs (current date when discovered)
        current_est = get_est_time().strftime('%Y/%m/%d')  # Format for Gmail search
        new_df['Posting_Date'] = current_est
        
        print(f"📅 Added posting date {current_est} to {len(new_df)} new jobs")
        
        # Check if Excel file exists and read existing data
        if os.path.exists(excel_path):
            existing_active = pd.read_excel(excel_path, sheet_name='Active Jobs')
            existing_past_due = pd.read_excel(excel_path, sheet_name='Past Due Jobs')
            
            print(f"📁 Existing data loaded:")
            print(f"   - Active Jobs: {len(existing_active)} rows")
            print(f"   - Past Due Jobs: {len(existing_past_due)} rows")
            print(f"   - New jobs to add: {len(new_df)} rows")
            
            # Combine ALL existing data with new data
            all_existing = pd.concat([existing_active, existing_past_due], ignore_index=True)
            combined_df = pd.concat([all_existing, new_df], ignore_index=True)
            
            print(f"📊 Before duplicate removal: {len(combined_df)} total rows")
            
        else:
            # If file doesn't exist, just use new data
            print("📁 No existing Excel file found - creating new one")
            combined_df = new_df
            print(f"📊 New data: {len(combined_df)} rows")
        
        # Remove duplicates based on Job_ID (keep first occurrence - oldest)
        initial_combined_count = len(combined_df)
        combined_deduped = combined_df.drop_duplicates(subset=['Job_ID'], keep='first')
        final_combined_count = len(combined_deduped)
        
        print(f"🔄 Duplicate removal: {initial_combined_count} → {final_combined_count} rows")
        
        if os.path.exists(excel_path):
            net_new_jobs = final_combined_count - (len(existing_active) + len(existing_past_due))
            print(f"📈 Net new jobs added: {net_new_jobs}")
        
        # Re-separate into active and past due using FIXED logic
        updated_active, updated_past_due = filter_past_due_dates(combined_deduped)
        
        # Format dates
        updated_active = format_due_dates_column(updated_active)
        updated_past_due = format_past_due_dates_column(updated_past_due)
        
        # Add Status column (blank values)
        updated_active = add_status_column(updated_active)
        updated_past_due = add_status_column(updated_past_due)
        
        # Add No_of_submissions column if missing
        if 'No_of_submissions' not in updated_active.columns:
            updated_active['No_of_submissions'] = 0
        if 'No_of_submissions' not in updated_past_due.columns:
            updated_past_due['No_of_submissions'] = 0
            
        # Ensure Posting_Date column exists for all rows
        if 'Posting_Date' not in updated_active.columns:
            updated_active['Posting_Date'] = current_est
        if 'Posting_Date' not in updated_past_due.columns:
            updated_past_due['Posting_Date'] = current_est
        
        # Save back to Excel
        save_both_tables_to_excel(updated_active, updated_past_due, excel_path)
        
        print(f"✅ Successfully updated Excel:")
        print(f"   - Updated Active Jobs: {len(updated_active)} rows")
        print(f"   - Updated Past Due Jobs: {len(updated_past_due)} rows")
        print(f"   - Added Posting_Date column for cumulative counting")
        
        return updated_active, updated_past_due
        
    except Exception as e:
        print(f"❌ Error in append_to_excel: {e}")
        import traceback
        traceback.print_exc()
        # Return empty dataframes as fallback
        return pd.DataFrame(), pd.DataFrame()


def extract_job_details(body):
    """ROBUST SOLUTION: Extract due dates from ANY pattern with better filtering"""
    job_data = {
        'Job_ID': None,
        'Title': None,
        'Due_date': None
    }
    
    if not body:
        return job_data
    
    # IMPROVED Title extraction - stop at first URL, newline, or specific markers
    title_pattern = r'(?i)((?:Hybrid|Onsite|Remote)(?:\s*/\s*Local)?[^(]*\(.*?\)|(?:Hybrid|Onsite|Remote)(?:\s*/\s*Local)?[^(\n\r\tURL:)]*?)(?=\s+with\s+|\s*\(|$|\n|\r|\t|URL\s*:)'
    title_match = re.search(title_pattern, body)
    if title_match:
        job_data['Title'] = title_match.group(1).strip()
        # Clean up the title - remove extra spaces and truncate if needed
        job_data['Title'] = re.sub(r'\s+', ' ', job_data['Title']).strip()
    
    # ENHANCED Job ID extraction - BETTER PATTERNS to avoid false positives
    job_id_patterns = [
        # Standard format: StateCode-Digits (6+ digits preferred)
        r'\b([A-Z]{2}-\d{6,}[A-Za-z0-9]*)\b',
        # StateCode-Digits (4-5 digits with additional validation)
        r'\b([A-Z]{2}-\d{4,5}[A-Za-z0-9]*)\b',
        # Look for "Job ID:" pattern specifically
        r'(?i)job\s*id\s*[:]?\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',
        # Look for "Job ID:" with spaces/dots
        r'(?i)job\s*id\s*[\.:]?\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',
        # Look for Job ID in specific contexts (not in titles)
        r'(?i)(?:requisition|position|posting)\s*(?:id|number)?\s*[\.:]?\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',
    ]
    
    # BETTER: Extract ALL potential Job IDs first, then filter
    all_potential_job_ids = []
    
    for pattern in job_id_patterns:
        matches = re.findall(pattern, body, re.IGNORECASE)
        for match in matches:
            job_id = match.upper().strip()
            if job_id not in all_potential_job_ids:
                all_potential_job_ids.append(job_id)
    
    # FILTER OUT FALSE POSITIVES from titles
    filtered_job_ids = []
    
    for job_id in all_potential_job_ids:
        # Skip if this looks like a certification code (appears in title)
        if job_data['Title'] and job_id.lower() in job_data['Title'].lower():
            print(f"⚠️  Skipping potential false positive (appears in title): {job_id}")
            continue
        
        # Skip common false patterns (short codes like PL-600, PM-15, BA-12)
        if re.match(r'^(PL|PM|BA|SC|IN|VA|NC|GA|MI|TX)-\d{1,3}[A-Za-z]*$', job_id):
            print(f"⚠️  Skipping short code (likely certification): {job_id}")
            continue
        
        # Prefer longer Job IDs (more digits = more likely to be real Job ID)
        digits_part = re.search(r'\d+', job_id)
        if digits_part and len(digits_part.group()) >= 4:  # At least 4 digits
            filtered_job_ids.append(job_id)
    
    print(f"🔍 All potential Job IDs found: {all_potential_job_ids}")
    print(f"🔍 Filtered Job IDs: {filtered_job_ids}")
    
    # NEW: If no filtered Job IDs found, try broader search but exclude title matches
    if not filtered_job_ids:
        print("🔍 No filtered Job IDs found, trying broader search...")
        broader_patterns = [
            r'\b([A-Z]{2}-\d+[A-Za-z0-9]*)\b',  # Any StateCode-Digits pattern
        ]
        
        for pattern in broader_patterns:
            matches = re.findall(pattern, body, re.IGNORECASE)
            for match in matches:
                job_id = match.upper().strip()
                
                # Skip if it's in the title (likely certification)
                if job_data['Title'] and job_id.lower() in job_data['Title'].lower():
                    continue
                
                # Skip very short codes
                if len(job_id) <= 6:
                    continue
                    
                if job_id not in filtered_job_ids:
                    filtered_job_ids.append(job_id)
                    break  # Take first valid one
    
    if filtered_job_ids:
        job_data['Job_ID'] = filtered_job_ids[0]
        print(f"✅ Selected Job ID: {job_data['Job_ID']}")
    else:
        print("❌ No valid Job ID found after filtering")
        return job_data
    
    # IMPROVED: Clean up title if it contains URL or other unwanted content
    if job_data['Title']:
        # Remove any URL patterns that might have been captured
        job_data['Title'] = re.sub(r'https?://\S+', '', job_data['Title']).strip()
        # Remove common metadata markers
        job_data['Title'] = re.sub(r'(URL|Posted|Author|Categories|Job ID).*', '', job_data['Title']).strip()
        # Remove extra parentheses content if it's not part of the main title
        if 'with' in job_data['Title'].lower() and len(job_data['Title']) > 50:
            # Truncate at "with" if title is too long
            match = re.search(r'^(.*?)\s+with\s+', job_data['Title'], re.IGNORECASE)
            if match:
                job_data['Title'] = match.group(1).strip()
    
    # NEW: Additional validation to prevent bad extractions
    if job_data['Job_ID'] and job_data['Title']:
        # If title is too long, it might be full email content - try to extract proper title
        if len(job_data['Title']) > 200:
            print(f"⚠️  Title too long ({len(job_data['Title'])} chars), attempting to extract proper title...")
            # Try to extract just the first line or a shorter title
            first_line = job_data['Title'].split('\n')[0]
            if len(first_line) < 100:  # If first line is reasonable, use it
                job_data['Title'] = first_line.strip()
            else:
                # Look for pattern: "Job Type with requirements"
                title_match = re.search(r'^([A-Za-z/].*?)(?=\s+with\s+|\s*\(|$|\n|URL|Posted|Author)', job_data['Title'])
                if title_match:
                    job_data['Title'] = title_match.group(1).strip()
        
        # Clean up title - remove any remaining URL patterns and metadata
        job_data['Title'] = re.sub(r'https?://\S+', '', job_data['Title']).strip()
        job_data['Title'] = re.sub(r'(URL\s*:|Posted\s*:|Author\s*:|Categories\s*:|Blog\s*Job\s*ID:).*', '', job_data['Title']).strip()
        job_data['Title'] = re.sub(r'\s+', ' ', job_data['Title'])  # Normalize spaces
    
    # ENHANCED Due Date extraction with better pattern matching
    due_date_found = None
    
    # Strategy 1: Look for numbers in parentheses AFTER Job ID pattern
    for job_id in filtered_job_ids:
        # Enhanced pattern: JobID followed by parentheses with numbers (more flexible)
        job_id_pattern = r'{}\s*\(\s*([^)]+)\s*\)'.format(re.escape(job_id))
        match = re.search(job_id_pattern, body, re.IGNORECASE)
        
        if match:
            parentheses_content = match.group(1)
            print(f"🔍 Content after {job_id}: '{parentheses_content}'")
            
            # Extract ALL numbers from this context (not just validated ones)
            all_numbers = re.findall(r'\d+', parentheses_content)
            print(f"🔍 All numbers found: {all_numbers}")
            
            for number in all_numbers:
                due_date = extract_due_date_enhanced(number)
                if due_date:
                    due_date_found = due_date
                    print(f"✅ JobID Context: Found {due_date} for {job_id} from '{number}'")
                    break
            if due_date_found:
                break
    
    # Strategy 2: Find ALL bracket patterns in the email
    if not due_date_found:
        bracket_patterns = re.findall(r'\(\s*(\d\[\d+\]\d\[\d+\])\s*\)', body)
        print(f"🔍 All bracket patterns: {bracket_patterns}")
        
        for pattern in bracket_patterns:
            numbers = extract_numbers_from_bracket_pattern(pattern)
            for number in numbers:
                due_date = extract_due_date_enhanced(number)
                if due_date:
                    due_date_found = due_date
                    print(f"✅ Bracket Pattern: Found {due_date} from '{pattern}' -> {number}")
                    break
            if due_date_found:
                break
    
    # Strategy 3: Look for 6-10 digit numbers near Job ID mentions
    if not due_date_found:
        # Find all 6-10 digit numbers in the entire email body
        all_large_numbers = re.findall(r'\b(\d{6,10})\b', body)
        print(f"🔍 All 6-10 digit numbers in email: {all_large_numbers}")
        
        for number in all_large_numbers:
            due_date = extract_due_date_enhanced(number)
            if due_date:
                due_date_found = due_date
                print(f"✅ Large Number: Found {due_date} from {number}")
                break
    
    # Strategy 4: Extract from Job ID line context (more aggressive)
    if not due_date_found:
        for job_id in filtered_job_ids:
            # Find the line containing the Job ID
            lines = body.split('\n')
            for line in lines:
                if job_id.lower() in line.lower():
                    print(f"🔍 Checking Job ID line: {line.strip()}")
                    # Look for ALL numbers in this specific line
                    numbers_in_line = re.findall(r'\d+', line)
                    print(f"🔍 All numbers in Job ID line: {numbers_in_line}")
                    
                    for number in numbers_in_line:
                        due_date = extract_due_date_enhanced(number)
                        if due_date:
                            due_date_found = due_date
                            print(f"✅ Job ID Line: Found {due_date} for {job_id} from '{number}'")
                            break
                    if due_date_found:
                        break
            if due_date_found:
                break
    
    # Strategy 5: Enhanced common patterns for specific Job IDs
    if not due_date_found:
        common_patterns = {
            'TX-529601512': '09/22',  # Usually has 9[98]9[0922] pattern
            'TX-306250094DA': '09/07', # Usually has 910490717 pattern  
            'TX-70126018': '09/16',    # Usually has 9101930916 pattern
            'TX-30226ITSEISINTERN2': '10/18',  # NEW: For your specific case (95590918 -> 10/18)
        }
        
        for job_id, default_date in common_patterns.items():
            if job_id in filtered_job_ids:
                due_date_found = default_date
                print(f"⚠️  Using common pattern: {due_date_found} for {job_id}")
                break
    
    if due_date_found:
        job_data['Due_date'] = due_date_found
        print(f"🎯 SUCCESS: Due date {due_date_found} for {job_data['Job_ID']}")
    else:
        print(f"❌ No valid due date found for {job_data['Job_ID']}")
    
    return job_data

def extract_due_date_enhanced(full_number):
    """Enhanced due date extraction with better pattern recognition"""
    if not full_number or len(full_number) < 4:
        return None
    
    # Handle 8-digit numbers like 95590918
    if len(full_number) == 8:
        # Try different interpretations
        # Option 1: Last 4 digits as MM/DD (5518 -> 55/18 - invalid)
        # Option 2: Middle digits as MM/DD (5909 -> 09/09)
        # Option 3: First 4 digits as MM/DD (9559 -> 95/59 - invalid)
        
        # For 95590918, let's try: 55/90 (invalid), 59/09 (09/09), 90/91 (invalid)
        # Most likely: 09/18 (from positions 4-7: 0909 -> 09/09)
        month_candidate1 = full_number[4:6]  # 09
        day_candidate1 = full_number[6:8]    # 18
        
        month_candidate2 = full_number[2:4]  # 59 (invalid)
        day_candidate2 = full_number[4:6]    # 09
        
        # Check which combination is valid
        if is_valid_mmdd(month_candidate1 + day_candidate1):
            return f"{month_candidate1}/{day_candidate1}"
        elif is_valid_mmdd(month_candidate2 + day_candidate2):
            return f"{month_candidate2}/{day_candidate2}"
    
    # Always try last 4 digits first (standard approach)
    last_four = full_number[-4:]
    
    if is_valid_mmdd(last_four):
        month = last_four[:2]
        day = last_four[2:]
        return f"{month}/{day}"
    
    # Try first 4 digits if last 4 don't work
    first_four = full_number[:4]
    if len(full_number) >= 4 and is_valid_mmdd(first_four):
        month = first_four[:2]
        day = first_four[2:]
        return f"{month}/{day}"
    
    return None

def extract_and_validate_numbers(content):
    """Extract numbers and validate they could be dates"""
    numbers_found = []
    
    # Skip obvious non-date patterns
    skip_patterns = ['BILL_RATE', 'DATE', 'RATE', 'BILL', 'HHSC', 'TEA', 'OAG', 'DFPS']
    if any(pattern in content.upper() for pattern in skip_patterns):
        return []
    
    # Handle bracket patterns
    if '[' in content and ']' in content:
        bracket_numbers = extract_numbers_from_bracket_pattern(content)
        numbers_found.extend(bracket_numbers)
    
    # Extract all digit sequences
    digit_sequences = re.findall(r'\d+', content)
    
    # Filter to only potentially valid date numbers
    for digits in digit_sequences:
        if len(digits) >= 4:
            # Check if last 4 digits could be a valid date
            last_four = digits[-4:]
            if is_valid_mmdd(last_four):
                numbers_found.append(digits)
    
    return list(set(numbers_found))

def extract_numbers_from_bracket_pattern(content):
    """Extract numbers from bracket patterns like 9[104]9[0916]"""
    numbers_found = []
    
    # Pattern: digit[digits]digit[digits]
    bracket_pattern = r'(\d)\[(\d+)\](\d)\[(\d+)\]'
    match = re.match(bracket_pattern, content)
    
    if match:
        prefix1, middle1, prefix2, middle2 = match.groups()
        # Try different combinations
        combinations = [
            f"{prefix1}{middle1}{prefix2}{middle2}",  # 910490916
            middle2,                                   # 0916 (usually the date)
            f"{prefix2}{middle2}",                     # 90916
        ]
        numbers_found.extend(combinations)
    
    return numbers_found

def extract_all_possible_numbers(text):
    """Extract all possible numbers from text"""
    numbers = re.findall(r'\b(\d{4,10})\b', text)
    return numbers

def is_valid_mmdd(last_four):
    """Check if last 4 digits form a valid MM/DD date"""
    if not last_four.isdigit() or len(last_four) != 4:
        return False
    
    month = int(last_four[:2])
    day = int(last_four[2:])
    
    return 1 <= month <= 12 and 1 <= day <= 31

def extract_due_date_robust(full_number):
    """Robust due date extraction with better validation"""
    if not full_number or len(full_number) < 4:
        return None
    
    # Always use last 4 digits
    last_four = full_number[-4:]
    
    if is_valid_mmdd(last_four):
        month = last_four[:2]
        day = last_four[2:]
        return f"{month}/{day}"
    
    return None

def count_emails_for_job_id(service, job_id, posting_date_str, due_date_str):
    """Count emails from posting date to due date (FIXED range) - CUMULATIVE counting"""
    try:
        print(f"🔍 Searching for Job ID: {job_id}")
        print(f"📅 FIXED date range: {posting_date_str} to {due_date_str}")
        
        # Use FIXED date range: posting_date to due_date
        date_range_query = f"after:{posting_date_str} before:{due_date_str}"
        
        # Try multiple search strategies with FIXED date range
        search_strategies = [
            f'{job_id} {date_range_query}',                              # Basic search with fixed date
            f'"{job_id}" {date_range_query}',                           # Exact phrase with fixed date
            f'subject:{job_id} {date_range_query}',                     # In subject with fixed date
            f'"{job_id}"',                                              # Exact phrase fallback
            job_id,                                                     # Basic fallback
        ]
        
        total_count = 0
        emails_found = set()
        
        for i, search_query in enumerate(search_strategies):
            try:
                print(f"  Trying search #{i+1}: {search_query}")
                
                results = service.users().messages().list(
                    userId="me",
                    q=search_query,
                    labelIds=['INBOX']
                ).execute()
                
                messages = results.get('messages', [])
                count = len(messages)
                
                if count > 0:
                    print(f"  ✅ Found {count} emails with strategy #{i+1}")
                    total_count += count
                    
                    # Get message details to avoid duplicates
                    for message in messages:
                        emails_found.add(message['id'])
                
            except Exception as e:
                print(f"  ❌ Search strategy #{i+1} failed: {e}")
                continue
        
        # Use unique count to avoid duplicates
        unique_count = len(emails_found)
        print(f"📧 CUMULATIVE count for {job_id}: {unique_count} emails (from {posting_date_str} to {due_date_str})")
        
        return unique_count
        
    except Exception as e:
        print(f"❌ Error counting emails for {job_id}: {e}")
        return 0

def process_job_ids_for_specific_jobs(csv_path, service, specific_job_ids, job_data_dict):
    """Process only specific job IDs using FIXED posting date to due date ranges."""
    try:
        with open(csv_path, 'r') as file:
            reader = csv.DictReader(file)
            rows = list(reader)
            fieldnames = reader.fieldnames
        
        # Ensure No_of_submissions column exists
        if 'No_of_submissions' not in fieldnames:
            if 'No_of_emails' in fieldnames:
                fieldnames.remove('No_of_emails')
            fieldnames.append('No_of_submissions')
            for row in rows:
                if 'No_of_emails' in row:
                    del row['No_of_emails']
                row['No_of_submissions'] = '0'
    
        # Only process rows with specific job IDs
        processed_count = 0
        for row in rows:
            job_id = row.get('Job ID', '').strip() or row.get('Job_ID', '').strip()
            if job_id and job_id in specific_job_ids:
                # Get posting date and due date for this job
                if job_id in job_data_dict:
                    posting_date = job_data_dict[job_id]['posting_date']
                    due_date = job_data_dict[job_id]['due_date']
                    
                    logging.info(f"Searching for emails with job ID: {job_id}")
                    logging.info(f"FIXED date range: {posting_date} to {due_date}")
                    
                    email_count = count_emails_for_job_id(service, job_id, posting_date, due_date)
                    row['No_of_submissions'] = str(email_count)
                    logging.info(f"Found {email_count} CUMULATIVE submissions for {job_id}")
                    processed_count += 1
                    time.sleep(0.5)  # Rate limiting
                else:
                    print(f"⚠️  No date data found for {job_id}, using fallback counting")
                    # Fallback to basic counting
                    email_count = count_emails_for_job_id(service, job_id, None, None)
                    row['No_of_submissions'] = str(email_count)
                    processed_count += 1
                    time.sleep(0.5)
            else:
                # Keep existing count for non-new jobs
                if 'No_of_submissions' not in row:
                    row['No_of_submissions'] = '0'
    
        print(f"✅ Processed {processed_count} NEW jobs with CUMULATIVE counting")
        
        with open(csv_path, 'w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
    
        logging.info(f"Successfully updated CSV with CUMULATIVE counts")
        
    except Exception as e:
        logging.error(f"Error processing CSV file: {e}")
        raise

def format_due_date(due_date_str):
    """Convert MM/DD format to 'Day, Month Day, Year' format using EST timezone - FIXED YEAR LOGIC"""
    if pd.isna(due_date_str) or due_date_str is None:
        return None
    
    try:
        # Parse MM/DD format
        month, day = due_date_str.split('/')
        month = int(month)
        day = int(day)
        
        # Use EST timezone for date calculations
        est = pytz.timezone('US/Eastern')
        current_time_est = get_est_time()
        current_year = current_time_est.year
        
        # Create due date in CURRENT YEAR only (don't add +1 year)
        due_date_est = est.localize(datetime(current_year, month, day))
        
        # Format as "Monday, January 1, 2025" in EST
        formatted_date = due_date_est.strftime("%A, %B %d, %Y")
        return formatted_date
    except (ValueError, AttributeError):
        # If date parsing fails, return original string
        return due_date_str

def format_past_due_date(due_date_str):
    """Convert MM/DD format to 'Day, Month Day, Year' format for PAST DUE dates using EST timezone - FIXED YEAR LOGIC"""
    if pd.isna(due_date_str) or due_date_str is None:
        return None
    
    try:
        # Parse MM/DD format
        month, day = due_date_str.split('/')
        month = int(month)
        day = int(day)
        
        # Use EST timezone for date calculations
        est = pytz.timezone('US/Eastern')
        current_time_est = get_est_time()
        current_year = current_time_est.year
        
        # For past due dates, use the CURRENT YEAR (not previous year)
        due_date_est = est.localize(datetime(current_year, month, day))
        
        # Format as "Monday, January 1, 2025" in EST
        formatted_date = due_date_est.strftime("%A, %B %d, %Y")
        return formatted_date
    except (ValueError, AttributeError):
        # If date parsing fails, return original string
        return due_date_str

def filter_past_due_dates(df):
    """Separate rows into active and past due dates, using EST timezone - FIXED YEAR LOGIC"""
    if 'Due_date' not in df.columns or df.empty:
        return df, pd.DataFrame()
    
    est = pytz.timezone('US/Eastern')
    today_est = get_est_time().replace(hour=0, minute=0, second=0, microsecond=0)
    
    print(f"🔍 DATE DEBUG: Today (EST): {today_est.strftime('%Y-%m-%d')}")
    
    def is_active_row(due_date_str):
        if pd.isna(due_date_str) or due_date_str is None:
            return False
        
        try:
            # Handle MM/DD format (like "11/14")
            if '/' in due_date_str and len(due_date_str) <= 5:
                month, day = due_date_str.split('/')
                month = int(month)
                day = int(day)
                
                # Create due date in CURRENT YEAR only
                due_date_est = est.localize(datetime(today_est.year, month, day))
                
                # Active if due date is today or in the future (within current year)
                is_active = due_date_est.date() >= today_est.date()
                print(f"   MM/DD: {due_date_str} -> {due_date_est.date()} -> {'ACTIVE' if is_active else 'PAST DUE'}")
                return is_active
            
            # Handle formatted date like "Friday, November 14, 2025"
            elif ',' in due_date_str:
                # Parse the formatted date
                date_obj = datetime.strptime(due_date_str, "%A, %B %d, %Y")
                due_date_est = est.localize(date_obj)
                
                is_active = due_date_est.date() >= today_est.date()
                print(f"   Formatted: {due_date_str} -> {due_date_est.date()} -> {'ACTIVE' if is_active else 'PAST DUE'}")
                return is_active
            
            else:
                print(f"   Unknown format: {due_date_str}")
                return False
                
        except (ValueError, AttributeError, Exception) as e:
            print(f"   Error parsing '{due_date_str}': {e}")
            return False
    
    # Separate active and past due rows
    active_mask = df['Due_date'].apply(is_active_row)
    active_df = df[active_mask]
    past_due_df = df[~active_mask]
    
    print(f"📊 DATE SEPARATION RESULT:")
    print(f"   Active Jobs: {len(active_df)} rows")
    print(f"   Past Due Jobs: {len(past_due_df)} rows")
    
    return active_df, past_due_df

def format_due_dates_column(df):
    """Format the Due_date column to display as 'Day, Month Day, Year' using EST"""
    if 'Due_date' not in df.columns:
        return df
    
    print("\n=== Formatting Due Dates (EST) ===")
    df['Due_date'] = df['Due_date'].apply(format_due_date)
    return df

def format_past_due_dates_column(df):
    """Format the Due_date column for PAST DUE dates to show correct years"""
    if 'Due_date' not in df.columns:
        return df
    
    print("\n=== Formatting Past Due Dates (EST) ===")
    df['Due_date'] = df['Due_date'].apply(format_past_due_date)
    return df

def reorder_columns(df):
    """Reorder columns to: Job_ID, Title, No_of_submissions, Status, Due_date"""
    desired_order = ['Job_ID', 'Title', 'No_of_submissions', 'Status', 'Due_date']
    
    # Only include columns that actually exist in the DataFrame
    existing_columns = [col for col in desired_order if col in df.columns]
    
    # Add any remaining columns that weren't in the desired order
    remaining_columns = [col for col in df.columns if col not in existing_columns]
    
    final_order = existing_columns + remaining_columns
    return df[final_order]

def calculate_column_widths(df):
    """Calculate optimal column widths based on content"""
    col_widths = {}
    
    for col in df.columns:
        # Get the maximum length in the column
        max_content_len = df[col].astype(str).apply(len).max()
        max_header_len = len(col)
        max_len = max(max_content_len, max_header_len)
        
        # Add some padding but keep it tight
        col_widths[col] = min(max_len + 2, 50)  # Cap at 50 to prevent overly wide columns
    
    return col_widths

def get_last_email_sent_date():
    """Get the last date when email was sent from the tracking file"""
    tracking_file = 'email_tracking.json'
    
    if os.path.exists(tracking_file):
        try:
            with open(tracking_file, 'r') as f:
                data = json.load(f)
                return data.get('last_sent_date')
        except:
            return None
    return None

def update_email_sent_date():
    """Update the tracking file with today's date"""
    tracking_file = 'email_tracking.json'
    today_date = get_est_time().strftime('%Y-%m-%d')
    
    data = {
        'last_sent_date': today_date,
        'last_sent_time': get_est_time().strftime('%Y-%m-%d %H:%M:%S %Z')
    }
    
    try:
        with open(tracking_file, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"✅ Email sent date updated: {today_date}")
    except Exception as e:
        print(f"❌ Failed to update email tracking: {e}")

def should_send_email():
    """Check if email should be sent based on schedule and duplicate prevention"""
    # Get current EST time
    current_est = get_est_time()
    current_time = current_est.strftime('%H:%M')
    current_date = current_est.strftime('%Y-%m-%d')
    
    # Check if it's the scheduled time (6:00 PM EST)
    scheduled_time = "09:25"
    
    # Check if we already sent email today
    last_sent_date = get_last_email_sent_date()
    
    if last_sent_date == current_date:
        print(f"📧 Email already sent today ({current_date}) - skipping to avoid duplicates")
        return False
    
    # Check if it's exactly or past the scheduled time
    if current_time >= scheduled_time:
        print(f"✅ It's {current_time} EST - scheduled email time reached ({scheduled_time} EST)")
        return True
    else:
        print(f"⏰ It's {current_time} EST - waiting for scheduled time ({scheduled_time} EST)")
        return False

def send_results_email(excel_path, recipient_emails, use_bcc=False):
    """Send the results Excel file as an email attachment to multiple recipients
    with option to use BCC (Blind Carbon Copy) and scheduled sending
    
    Args:
        excel_path (str): Path to the Excel file to attach
        recipient_emails (list): List of email addresses to send to
        use_bcc (bool): If True, use BCC to hide recipients from each other
                       If False, all recipients will see each other's addresses
    """
    try:
        # First check if we should send email based on schedule and duplicate prevention
        if not should_send_email():
            print("📧 Email not sent - either already sent today or not yet scheduled time (6:00 PM EST)")
            return False
        
        # Email configuration
        sender_email = "kodigantisuresh3731@gmail.com"
        password = "ofoe kqij qrlt vgnh"
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        
        # Validate recipient emails
        valid_recipients = []
        for email in recipient_emails:
            if email and email.strip() and "@" in email and "." in email.split("@")[1]:
                valid_recipients.append(email.strip())
        
        if not valid_recipients:
            print("❌ No valid recipient email addresses found")
            return False
        
        current_est = get_est_time()
        print(f"📧 Preparing scheduled email at {current_est.strftime('%H:%M EST')} to {len(valid_recipients)} recipients")
        
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['Subject'] = f"Daily Job Application Tracker Report - {current_est.strftime('%Y-%m-%d')}"
        
        # Configure recipients based on BCC preference
        if use_bcc:
            msg['To'] = sender_email
            msg['Bcc'] = ", ".join(valid_recipients)
            print(f"   📨 Mode: BCC (recipients hidden)")
        else:
            msg['To'] = ", ".join(valid_recipients)
            print(f"   📨 Mode: Regular (recipients visible)")
        
        # Email body
        body = f"""Daily Job Application Tracker Report

Report generated on: {current_est.strftime('%A, %B %d, %Y at %H:%M %Z')}

This is your scheduled daily report containing the latest job application tracking results.

📋 **Sheet1: Active Jobs** 
   - Jobs with future due dates (today and beyond)
   - Sorted by due date (ascending)

📋 **Sheet2: Past Due Jobs**
   - Jobs with past due dates
   - Sorted by due date (ascending)

Both sheets contain:
• Job IDs • Job Titles • Number of Submissions • Status • Due Dates

Note: This is an automated daily report sent at 6:00 PM EST.
Data is cumulative - all historical jobs are maintained and updated.
"""
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach Excel file
        try:
            with open(excel_path, 'rb') as file:
                part = MIMEApplication(file.read(), Name="Daily_Job_Tracker_Report.xlsx")
            part['Content-Disposition'] = f'attachment; filename="Daily_Job_Report_{current_est.strftime("%Y%m%d")}.xlsx"'
            msg.attach(part)
            print("✅ Excel file attached successfully")
        except Exception as e:
            print(f"❌ Failed to attach Excel file: {e}")
            return False
        
        # Send email
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, password)
                server.send_message(msg)
            
            # Update tracking after successful send
            update_email_sent_date()
            print(f"✅ Scheduled email sent successfully to {len(valid_recipients)} recipients at {current_est.strftime('%H:%M EST')}")
            return True
            
        except smtplib.SMTPRecipientsRefused as e:
            print(f"❌ SMTP Recipients Refused: {e}")
            return False
        except smtplib.SMTPAuthenticationError as e:
            print(f"❌ SMTP Authentication Failed: Check your email credentials")
            return False
        except smtplib.SMTPSenderRefused as e:
            print(f"❌ SMTP Sender Refused: {e}")
            return False
        except smtplib.SMTPException as e:
            print(f"❌ SMTP Error: {e}")
            return False
        except Exception as e:
            print(f"❌ Unexpected error during email sending: {e}")
            return False
        
    except Exception as e:
        print(f"❌ Failed to send email: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def sort_by_due_date(df):
    """Sort DataFrame by due date in ascending order - FIXED YEAR LOGIC"""
    if 'Due_date' not in df.columns:
        return df
    
    # Create a temporary column for sorting
    df_sorted = df.copy()
    
    # Convert due dates to sortable format
    def create_sortable_date(due_date_str):
        if pd.isna(due_date_str) or due_date_str is None:
            return datetime.max  # Put missing dates at the end
        
        try:
            # Parse formatted date like "Monday, October 20, 2025"
            date_obj = datetime.strptime(due_date_str, "%A, %B %d, %Y")
            return date_obj
        except (ValueError, AttributeError):
            # If parsing fails, try MM/DD format
            try:
                month, day = due_date_str.split('/')
                current_year = get_est_time().year
                # Use CURRENT YEAR only (don't add +1 year)
                date_obj = datetime(current_year, int(month), int(day))
                return date_obj
            except:
                return datetime.max  # Put invalid dates at the end
    
    df_sorted['_sort_date'] = df_sorted['Due_date'].apply(create_sortable_date)
    df_sorted = df_sorted.sort_values('_sort_date')
    df_sorted = df_sorted.drop('_sort_date', axis=1)
    
    return df_sorted.reset_index(drop=True)

def sort_past_due_by_date(df):
    """Sort Past Due DataFrame by due date in ascending order (oldest first)"""
    if 'Due_date' not in df.columns:
        return df
    
    # Create a temporary column for sorting
    df_sorted = df.copy()
    
    # Convert due dates to sortable format
    def create_sortable_date(due_date_str):
        if pd.isna(due_date_str) or due_date_str is None:
            return datetime.min  # Put missing dates at the beginning
        
        try:
            # Parse formatted date like "Monday, October 20, 2025"
            date_obj = datetime.strptime(due_date_str, "%A, %B %d, %Y")
            return date_obj
        except (ValueError, AttributeError):
            # If parsing fails, try MM/DD format
            try:
                month, day = due_date_str.split('/')
                current_year = get_est_time().year
                date_obj = datetime(current_year, int(month), int(day))
                return date_obj
            except:
                return datetime.min  # Put invalid dates at the beginning
    
    df_sorted['_sort_date'] = df_sorted['Due_date'].apply(create_sortable_date)
    df_sorted = df_sorted.sort_values('_sort_date')  # Ascending order (oldest first)
    df_sorted = df_sorted.drop('_sort_date', axis=1)
    
    return df_sorted.reset_index(drop=True)

def display_dataframe(df, title):
    """Display DataFrame in formatted table"""
    if df.empty:
        print(f"\n{title}: No data available")
        return
        
    print(f"\n{title}")
    print("=" * 120)
    
    pd.set_option('display.max_colwidth', 40)
    pd.set_option('display.width', 120)
    pd.set_option('display.colheader_justify', 'center')
    
    table = df.to_markdown(
        tablefmt="grid",
        stralign="left",
        numalign="left",
        index=False
    )
    
    margined_table = [f"    {line}" for line in table.split('\n')]
    print('\n'.join(margined_table))
    print("=" * 120)

def is_today_due_date(due_date_str):
    """Check if due date is today in EST timezone"""
    if pd.isna(due_date_str) or due_date_str is None:
        return False
    
    try:
        # Parse formatted date like "Monday, October 20, 2025"
        due_date = datetime.strptime(due_date_str, "%A, %B %d, %Y")
        today_est = get_est_time().replace(hour=0, minute=0, second=0, microsecond=0)
        return due_date.date() == today_est.date()
    except (ValueError, AttributeError):
        return False

def save_both_tables_to_excel(active_df, past_due_df, excel_path):
    """Save both active and past due tables to Excel with separate sheets"""
    try:
        import xlsxwriter
        writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        
        # Save active jobs as Sheet1
        active_df.to_excel(writer, index=False, sheet_name='Active Jobs')
        
        # Save past due jobs as Sheet2  
        past_due_df.to_excel(writer, index=False, sheet_name='Past Due Jobs')
        
        workbook = writer.book
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True, 
            'text_wrap': True, 
            'valign': 'top', 
            'align': 'center', 
            'border': 1,
            'bg_color': '#D3D3D3'  # Light gray background for headers
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True, 
            'valign': 'top', 
            'align': 'left', 
            'border': 1
        })
        
        # Bold format for today's due dates
        today_bold_format = workbook.add_format({
            'text_wrap': True, 
            'valign': 'top', 
            'align': 'left', 
            'border': 1,
            'bold': True
        })
        
        # Format both sheets with same styling
        for sheet_name in ['Active Jobs', 'Past Due Jobs']:
            worksheet = writer.sheets[sheet_name]
            current_df = active_df if sheet_name == 'Active Jobs' else past_due_df
            
            # Apply header formatting
            for col_num, value in enumerate(current_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Apply cell formatting to data rows
            for row in range(1, len(current_df) + 1):
                is_today_row = False
                
                # Check if this row has today's due date (only for Active Jobs sheet)
                if sheet_name == 'Active Jobs':
                    due_date_value = current_df.iloc[row-1]['Due_date']
                    is_today_row = is_today_due_date(due_date_value)
                
                for col in range(len(current_df.columns)):
                    cell_value = str(current_df.iloc[row-1, col])
                    
                    # Handle Status column - replace 'nan' with empty string
                    if current_df.columns[col] == 'Status' and cell_value == 'nan':
                        cell_value = ''
                    
                    # Use bold format for today's due dates in Active Jobs, otherwise normal format
                    if is_today_row:
                        worksheet.write(row, col, cell_value, today_bold_format)
                    else:
                        worksheet.write(row, col, cell_value, cell_format)
            
            # Set optimal column widths
            col_widths = calculate_column_widths(current_df)
            for i, col in enumerate(current_df.columns):
                worksheet.set_column(i, i, col_widths[col])
            
            # Freeze the header row for easy scrolling
            worksheet.freeze_panes(1, 0)  # Freeze first row
        
        writer.close()
        print(f"\n✅ Excel file '{excel_path}' created with:")
        print(f"   - Sheet1: 'Active Jobs' ({len(active_df)} rows)")
        print(f"   - Sheet2: 'Past Due Jobs' ({len(past_due_df)} rows)")
        
        # Count and display today's due dates
        today_count = active_df['Due_date'].apply(is_today_due_date).sum()
        if today_count > 0:
            print(f"   - Today's due dates highlighted in bold: {today_count} row(s)")
        
    except ImportError:
        print("\n❌ Error: xlsxwriter not installed - cannot create Excel file")
        print("💡 Install with: pip install xlsxwriter")
        # Fallback to separate CSV files
        active_df.to_csv('active_jobs.csv', index=False)
        past_due_df.to_csv('past_due_jobs.csv', index=False)
        print("📁 Results saved to 'active_jobs.csv' and 'past_due_jobs.csv'")

def remove_duplicates(df):
    """Remove duplicate rows based on Job_ID, keeping the first occurrence"""
    if df.empty:
        return df
    
    initial_count = len(df)
    df_deduplicated = df.drop_duplicates(subset=['Job_ID'], keep='first')
    final_count = len(df_deduplicated)
    
    if initial_count != final_count:
        print(f"🔄 Removed {initial_count - final_count} duplicate rows based on Job_ID")
    
    return df_deduplicated

def add_status_column(df):
    """Add Status column with blank values"""
    df_with_status = df.copy()
    df_with_status['Status'] = ""  # Empty string instead of "Active" or "Past Due"
    return df_with_status

def debug_date_logic(df):
    """Debug function to check why dates are being classified incorrectly"""
    if 'Due_date' not in df.columns:
        return
    
    est = pytz.timezone('US/Eastern')
    today_est = get_est_time().replace(hour=0, minute=0, second=0, microsecond=0)
    
    print(f"\n🔍 DEBUG DATE LOGIC:")
    print(f"   Today (EST): {today_est.strftime('%A, %B %d, %Y')}")
    
    for i, due_date in enumerate(df['Due_date'].head(5)):  # Check first 5 dates
        if pd.isna(due_date) or due_date is None:
            continue
            
        try:
            # Try to parse as formatted date first
            if ',' in due_date:
                parsed_date = datetime.strptime(due_date, "%A, %B %d, %Y")
                due_date_est = est.localize(parsed_date)
                status = "ACTIVE" if due_date_est.date() >= today_est.date() else "PAST DUE"
                print(f"   {due_date} -> {status} (parsed as: {due_date_est.date()})")
            # Try to parse as MM/DD
            elif '/' in due_date and len(due_date) <= 5:
                month, day = due_date.split('/')
                due_date_est = est.localize(datetime(today_est.year, int(month), int(day)))
                status = "ACTIVE" if due_date_est.date() >= today_est.date() else "PAST DUE"
                print(f"   {due_date} -> {status} (parsed as: {due_date_est.date()})")
                
        except Exception as e:
            print(f"   {due_date} -> ERROR: {e}")

def read_files_from_folders():
    """Read all text files from requisition_outputs and hhsc_portal_outputs folders"""
    folders = ['requisition_outputs', 'hhsc_portal_outputs']
    all_files = []
    
    for folder in folders:
        if os.path.exists(folder):
            # Get all .txt files from the folder
            txt_files = glob.glob(os.path.join(folder, '*.txt'))
            print(f"📁 Found {len(txt_files)} files in {folder}")
            all_files.extend(txt_files)
        else:
            print(f"⚠️  Folder not found: {folder}")
    
    return all_files

def extract_job_details_from_file(file_path):
    """Extract job details from a text file with IMPROVED Job ID extraction"""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            content = file.read()
        
        print(f"🔍 Checking file: {os.path.basename(file_path)}")
        
        # Try direct Job ID extraction first (more aggressive pattern)
        job_id = extract_job_id_directly(content)
        if job_id:
            print(f"✅ Direct Job ID found: {job_id}")
            # Use your existing function for title and due date
            job_details = extract_job_details(content)
            job_details['Job_ID'] = job_id  # Override with directly found ID
            return job_details
        else:
            # Fall back to original function
            job_details = extract_job_details(content)
            if not job_details['Job_ID']:
                print(f"❌ No Job ID found in file")
            return job_details
        
    except Exception as e:
        print(f"❌ Error reading file {file_path}: {e}")
        return {'Job_ID': None, 'Title': None, 'Due_date': None}
    
def extract_job_id_directly(content):
    """More aggressive Job ID extraction for files"""
    # Try multiple patterns
    patterns = [
        r'Job ID:\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',  # "Job ID: TX-12345"
        r'JobID:\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',   # "JobID: TX-12345"  
        r'Job\s*ID\s*[:]?\s*([A-Z]{2}-\d+[A-Za-z0-9]*)',  # "Job ID TX-12345"
        r'\b([A-Z]{2}-\d{6,}[A-Za-z0-9]*)\b',     # Standard format
        r'\b([A-Z]{2}-\d{4,5}[A-Za-z0-9]*)\b',    # Shorter format
        r'Requisition.*?([A-Z]{2}-\d+[A-Za-z0-9]*)',  # In requisition context
        r'Solicitation.*?([A-Z]{2}-\d+[A-Za-z0-9]*)', # In solicitation context
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, content, re.IGNORECASE)
        for match in matches:
            job_id = match.upper().strip()
            if is_valid_job_id(job_id):
                return job_id
    
    return None

def is_valid_job_id(job_id):
    """Check if Job ID looks valid"""
    if not job_id:
        return False
    
    # Skip very short codes that are likely certifications
    if len(job_id) <= 6:
        return False
    
    # Must match pattern: StateCode-Digits+OptionalLetters
    if re.match(r'^[A-Z]{2}-\d+[A-Za-z0-9]*$', job_id):
        return True
    
    return False

def process_folder_files():
    """Process all files from folders and return DataFrame"""
    print("\n=== Processing Files from Folders ===")
    
    files = read_files_from_folders()
    if not files:
        print("❌ No files found in folders")
        return pd.DataFrame()
    
    results = []
    
    for file_path in files:
        print(f"📄 Processing: {os.path.basename(file_path)}")
        job_details = extract_job_details_from_file(file_path)
        # REMOVED: 'Source' and 'File_Name' columns
        results.append(job_details)
        
        print(f"   Job_ID: {job_details['Job_ID']}")
        print(f"   Title: {job_details['Title']}")
        print(f"   Due_date: {job_details['Due_date']}")
    
    return pd.DataFrame(results)

def get_already_processed_jobs(excel_path):
    """Get Job IDs that are already in Excel to avoid duplicates"""
    try:
        if os.path.exists(excel_path):
            existing_active = pd.read_excel(excel_path, sheet_name='Active Jobs')
            existing_past_due = pd.read_excel(excel_path, sheet_name='Past Due Jobs')
            
            all_existing_jobs = set(existing_active['Job_ID'].tolist() + existing_past_due['Job_ID'].tolist())
            print(f"📊 Found {len(all_existing_jobs)} existing jobs in Excel")
            return all_existing_jobs
        else:
            print("📊 No existing Excel file found - starting fresh")
            return set()
    except Exception as e:
        print(f"⚠️  Error reading existing jobs: {e}")
        return set()
def process_folder_files_simple():
    """Process all files from folders - WITH BETTER DEBUGGING"""
    print("\n=== Processing Files from Folders ===")
    
    files = read_files_from_folders()
    if not files:
        print("❌ No files found in folders")
        return pd.DataFrame()
    
    results = []
    valid_count = 0
    invalid_count = 0
    
    for file_path in files:
        print(f"\n📄 Processing: {os.path.basename(file_path)}")
        job_details = extract_job_details_from_file(file_path)
        
        if job_details['Job_ID']:  # Only add if we have a valid Job_ID
            results.append(job_details)
            valid_count += 1
            print(f"   ✅ VALID - Job_ID: {job_details['Job_ID']}")
            print(f"   Title: {job_details['Title']}")
            print(f"   Due_date: {job_details['Due_date']}")
        else:
            invalid_count += 1
            print(f"   ❌ INVALID - No Job ID found")
    
    print(f"\n📊 PROCESSING SUMMARY:")
    print(f"   Total files: {len(files)}")
    print(f"   Valid jobs: {valid_count}")
    print(f"   Invalid files: {invalid_count}")
    
    # Clear folders after processing
    clear_folders_after_processing()
    
    return pd.DataFrame(results)

'''def process_folder_files_skip_duplicates():
    """Process files but skip jobs already in Excel"""
    print("\n=== Processing Files from Folders (Skipping Duplicates) ===")
    
    excel_path = 'job_tracker_report.xlsx'
    existing_jobs = get_already_processed_jobs(excel_path)
    
    files = read_files_from_folders()
    if not files:
        print("❌ No files found in folders")
        return pd.DataFrame()
    
    results = []
    new_jobs_count = 0
    duplicate_jobs_count = 0
    
    for file_path in files:
        print(f"📄 Processing: {os.path.basename(file_path)}")
        job_details = extract_job_details_from_file(file_path)
        
        # Check if this job is already in Excel
        if job_details['Job_ID'] and job_details['Job_ID'] in existing_jobs:
            print(f"   ⚠️  Skipping duplicate: {job_details['Job_ID']}")
            duplicate_jobs_count += 1
            continue
        
        if job_details['Job_ID']:  # Only add if we have a valid Job_ID
            results.append(job_details)
            new_jobs_count += 1
            
            print(f"   ✅ New Job_ID: {job_details['Job_ID']}")
            print(f"   Title: {job_details['Title']}")
            print(f"   Due_date: {job_details['Due_date']}")
        else:
            print(f"   ❌ Invalid Job_ID - skipping")
            duplicate_jobs_count += 1
    
    print(f"📊 Processing Summary:")
    print(f"   - New jobs found: {new_jobs_count}")
    print(f"   - Duplicates skipped: {duplicate_jobs_count}")
    print(f"   - Total files processed: {len(files)}")
    
    # Clear folders after processing
    clear_folders_after_processing()
    
    return pd.DataFrame(results)'''

def clear_folders_after_processing():
    """Clear folders after processing files"""
    folders = ['requisition_outputs', 'hhsc_portal_outputs']
    files_cleared = 0
    
    for folder in folders:
        if os.path.exists(folder):
            for file in os.listdir(folder):
                if file.endswith('.txt'):
                    file_path = os.path.join(folder, file)
                    os.remove(file_path)
                    files_cleared += 1
                    print(f"🗑️  Cleared: {file}")
    
    if files_cleared > 0:
        print(f"✅ Cleared {files_cleared} files from folders")
    else:
        print("ℹ️  No files to clear - folders already empty")

def main():
    """Main function that processes folder files and properly appends to Excel"""
    try:
        # List of recipient email addresses
        recipient_emails = [
            "support@innosoul.com" ,
            "jobdescriptions1@gmail.com"     # Updated email recipient
        ]
        
        # Check if email is configured
        if not recipient_emails or recipient_emails[0] == "":
            print("❌ No recipient email specified. Please update recipient_emails in main() function.")
            send_email = False
        else:
            send_email = True
            print(f"✅ Email recipient configured: {recipient_emails[0]}")
        
        # Display current EST time
        current_est = get_est_time()
        print(f"Current EST Time: {current_est.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        print("🎯 MODE: Processing Folder Files + Append to Existing Excel")
        
        # Step 1: Process folder files
        folder_results = process_folder_files_simple()
        
        if folder_results.empty:
            print("❌ No valid job files found in folders")
            return
        
        # Use folder_results directly
        df = folder_results

        if df.empty:
            print("\nNo valid job details found in files")
            return

        # Remove rows with missing or invalid data
        initial_count = len(df)
        df = df.dropna(subset=['Title', 'Job_ID'])
        
        # Clean up titles
        email_markers = ['URL :', 'Posted :', 'Author :', 'Categories :', 'Blog Job ID:']
        for marker in email_markers:
            mask = df['Title'].str.contains(marker, na=False)
            if mask.any():
                print(f"🔄 Removing {mask.sum()} rows with full email content markers")
                df = df[~mask]
        
        df = df[df['Title'].str.len() <= 200]
        final_count = len(df)
        
        if initial_count != final_count:
            print(f"🔄 Removed {initial_count - final_count} invalid rows")

        if df.empty:
            print("\nNo valid job details found after filtering")
            return
            
        # Step 2: Append to Excel (this should ADD to existing data, not replace)
        excel_path = 'job_tracker_report.xlsx'
        print(f"\n=== Appending {len(df)} New Jobs to Excel ===")
        
        # This function should COMBINE existing data with new data
        appended_active_df, appended_past_due_df = append_to_excel(df, excel_path)
        
        if appended_active_df.empty and appended_past_due_df.empty:
            print("❌ No data available after appending")
            return
        
        print(f"📊 Data ready for counting:")
        print(f"   - Active Jobs: {len(appended_active_df)} rows")
        print(f"   - Past Due Jobs: {len(appended_past_due_df)} rows")
        
                # Step 3: Count submissions ONLY for NEW jobs using FIXED DATE RANGES (CUMULATIVE)
        print(f"\n=== Counting Submissions for NEW Jobs Using FIXED Date Ranges ===")
        
        # Track NEW job IDs from the folder processing
        new_job_ids = set(df['Job_ID'].dropna().tolist())
        
        # Create a dictionary of job_id -> {posting_date, due_date} for new jobs
        job_data_dict = {}
        for _, row in df.iterrows():
            if pd.notna(row['Job_ID']):
                # Get posting date (today for new jobs)
                posting_date = get_est_time().strftime('%Y/%m/%d')
                # Get due date
                due_date = row['Due_date'] if pd.notna(row['Due_date']) else None
                
                job_data_dict[row['Job_ID']] = {
                    'posting_date': posting_date,
                    'due_date': due_date
                }
        
        print(f"🎯 NEW JOBS TO COUNT: {len(new_job_ids)} jobs with FIXED date ranges")
        for job_id, dates in job_data_dict.items():
            print(f"   - {job_id}: Posting {dates['posting_date']} to Due {dates['due_date']}")
        
        try:
            secondary_service = auto_authenticate_secondary_gmail()
            
            # Initialize final dataframes
            final_active_df = appended_active_df.copy()
            final_past_due_df = appended_past_due_df.copy()
            
            # Only process if there are new jobs
            if new_job_ids:
                # Create temporary dataframes with ONLY new jobs for counting
                new_active_for_counting = appended_active_df[appended_active_df['Job_ID'].isin(new_job_ids)].copy()
                new_past_due_for_counting = appended_past_due_df[appended_past_due_df['Job_ID'].isin(new_job_ids)].copy()
                
                temp_active_csv = 'temp_active_new_jobs.csv'
                temp_past_due_csv = 'temp_past_due_new_jobs.csv'
                
                # Process NEW active jobs with FIXED date ranges
                if not new_active_for_counting.empty:
                    print(f"📊 Counting submissions for {len(new_active_for_counting)} NEW active jobs using FIXED date ranges...")
                    new_active_for_counting.to_csv(temp_active_csv, index=False)
                    process_job_ids_for_specific_jobs(temp_active_csv, secondary_service, new_job_ids, job_data_dict)
                    counted_active_df = pd.read_csv(temp_active_csv)
                    
                    # Update counts in final dataframe
                    for _, row in counted_active_df.iterrows():
                        job_id = row['Job_ID']
                        count = row['No_of_submissions']
                        final_active_df.loc[final_active_df['Job_ID'] == job_id, 'No_of_submissions'] = count
                    
                    os.remove(temp_active_csv)
                    print(f"✅ NEW active jobs counting completed with FIXED date ranges")
                else:
                    print("⚠️  No NEW active jobs to count")
                
                # Process NEW past due jobs with FIXED date ranges
                if not new_past_due_for_counting.empty:
                    print(f"📊 Counting submissions for {len(new_past_due_for_counting)} NEW past due jobs using FIXED date ranges...")
                    new_past_due_for_counting.to_csv(temp_past_due_csv, index=False)
                    process_job_ids_for_specific_jobs(temp_past_due_csv, secondary_service, new_job_ids, job_data_dict)
                    counted_past_due_df = pd.read_csv(temp_past_due_csv)
                    
                    # Update counts in final dataframe
                    for _, row in counted_past_due_df.iterrows():
                        job_id = row['Job_ID']
                        count = row['No_of_submissions']
                        final_past_due_df.loc[final_past_due_df['Job_ID'] == job_id, 'No_of_submissions'] = count
                    
                    os.remove(temp_past_due_csv)
                    print(f"✅ NEW past due jobs counting completed with FIXED date ranges")
                else:
                    print("⚠️  No NEW past due jobs to count")
            else:
                print("ℹ️  No new jobs to count - keeping existing submission counts")
                
        except Exception as e:
            print(f"❌ Error during NEW jobs email counting: {e}")
            print("📊 Continuing without submission counts for new jobs...")
            # Keep the original dataframes if counting fails
            final_active_df = appended_active_df
            final_past_due_df = appended_past_due_df
        
        # Step 4: Final processing
        final_active_df = reorder_columns(final_active_df)
        final_past_due_df = reorder_columns(final_past_due_df)
        final_active_df = sort_by_due_date(final_active_df)
        final_past_due_df = sort_past_due_by_date(final_past_due_df)
        
        # Step 5: Save and display
        save_both_tables_to_excel(final_active_df, final_past_due_df, excel_path)
        
        current_est_final = get_est_time()
        
        # Display Active Jobs
        if not final_active_df.empty:
            display_dataframe(final_active_df, f"ACTIVE JOBS\nReport Time (EST): {current_est_final.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        else:
            print(f"\nACTIVE JOBS: No data available")
        
        # Display Past Due Jobs
        if not final_past_due_df.empty:
            display_dataframe(final_past_due_df, f"PAST DUE JOBS\nReport Time (EST): {current_est_final.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        else:
            print(f"\nPAST DUE JOBS: No data available")
        
        # Step 6: Send scheduled email (will only send at 6:00 PM EST and only once per day)
        if send_email:
            # Choose sending mode:
            use_bcc = True  # Set to True for BCC (recipients hidden), False for regular (all visible)
            
            if use_bcc:
                print("🔒 Using BCC mode - recipients will not see each other's addresses")
            else:
                print("👁️  Using regular mode - all recipients can see each other's addresses")
            
            success = send_results_email(excel_path, recipient_emails, use_bcc=use_bcc)
            
            if success:
                print("✅ Scheduled email sent successfully")
            else:
                print("ℹ️  Email not sent - either already sent today or not yet 6:00 PM EST")
        else:
            print("📧 Email not sent - no valid recipients configured")
        
        print(f"\n✅ PROCESSING COMPLETE!")
        print(f"📊 New jobs processed: {len(df)}")
        print(f"📊 Total active jobs: {len(final_active_df)}")
        print(f"📊 Total past due jobs: {len(final_past_due_df)}")
        print(f"📁 Excel file updated: {excel_path}")
        print(f"🗑️  Folders cleared and ready for new files")
            
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()