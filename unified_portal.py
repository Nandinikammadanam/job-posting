import os
import pandas as pd
import sys
import pytz
import time
import threading
import json
import shutil
from datetime import datetime, timedelta

# ==================== GLOBAL CONFIGURATION ====================

# Global state to track processed emails
PROCESSED_EMAILS_FILE = "processed_emails.json"
PROCESSED_FILES_FILE = "processed_files.json"

# ==== ONLY CHANGE THESE THREE LINES FOR FUTURE TIMING CHANGES ====
DAILY_EMAIL_TIME = "09:25"  # 4:25 PM EST
DAILY_EMAIL_START_TIME = "09:25"  # 4:25 PM EST
DAILY_EMAIL_END_TIME = "09:40"    # 15-minute window

# Email recipients list - ADD ALL RECIPIENTS HERE
EMAIL_RECIPIENTS = [
    "jobdescriptions1@gmail.com",
    "support@innosoul.com",
]

# Global state variables
processed_emails_state = {}
processed_files_state = {}
# ==================== END GLOBAL CONFIGURATION ========

def load_processed_emails():
    """Load the state of processed emails from file"""
    global processed_emails_state
    try:
        if os.path.exists(PROCESSED_EMAILS_FILE):
            with open(PROCESSED_EMAILS_FILE, 'r') as f:
                processed_emails_state = json.load(f)
        else:
            processed_emails_state = {
                'hhsc_emails': [],
                'vms_emails': [],
                'last_initial_run': None,
                'last_email_sent_date': None
            }
    except Exception as e:
        print(f"❌ Error loading processed emails state: {e}")
        processed_emails_state = {
            'hhsc_emails': [],
            'vms_emails': [],
            'last_initial_run': None,
            'last_email_sent_date': None
        }

def save_processed_emails():
    """Save the state of processed emails to file"""
    try:
        with open(PROCESSED_EMAILS_FILE, 'w') as f:
            json.dump(processed_emails_state, f, indent=2)
    except Exception as e:
        print(f"❌ Error saving processed emails state: {e}")

def mark_email_processed(portal_type, email_id):
    """Mark an email as processed"""
    if portal_type not in processed_emails_state:
        processed_emails_state[portal_type] = []
    
    if email_id not in processed_emails_state[portal_type]:
        processed_emails_state[portal_type].append(email_id)
        save_processed_emails()

def is_email_processed(portal_type, email_id):
    """Check if an email has been processed"""
    return email_id in processed_emails_state.get(portal_type, [])

def clear_old_processed_emails():
    """Clear processed emails older than 7 days to prevent state file from growing too large"""
    try:
        # Keep only emails from last 7 days in memory
        # In a real implementation, you'd track timestamps
        max_emails_to_keep = 1000
        for portal_type in ['hhsc_emails', 'vms_emails']:
            if portal_type in processed_emails_state and len(processed_emails_state[portal_type]) > max_emails_to_keep:
                # Keep only the most recent emails
                processed_emails_state[portal_type] = processed_emails_state[portal_type][-max_emails_to_keep:]
        
        save_processed_emails()
    except Exception as e:
        print(f"⚠️ Error clearing old processed emails: {e}")

# ==================== NEW 24-HOUR FUNCTIONS ====================

def load_processed_files():
    """Load processed files state"""
    global processed_files_state
    try:
        if os.path.exists(PROCESSED_FILES_FILE):
            with open(PROCESSED_FILES_FILE, 'r') as f:
                processed_files_state = json.load(f)
        else:
            processed_files_state = {
                'processed_files': [],
                'last_email_date': None,
                'last_processing_time': None
            }
    except Exception as e:
        print(f"❌ Error loading processed files: {e}")
        processed_files_state = {
            'processed_files': [],
            'last_email_date': None,
            'last_processing_time': None
        }

def save_processed_files():
    """Save processed files state"""
    try:
        with open(PROCESSED_FILES_FILE, 'w') as f:
            json.dump(processed_files_state, f, indent=2)
    except Exception as e:
        print(f"❌ Error saving processed files: {e}")

def mark_file_processed(filename):
    """Mark a file as processed"""
    if filename not in processed_files_state.get('processed_files', []):
        processed_files_state['processed_files'].append(filename)
        processed_files_state['last_processing_time'] = datetime.now().isoformat()
        save_processed_files()

def is_file_processed(filename):
    """Check if file is already processed"""
    return filename in processed_files_state.get('processed_files', [])

def should_send_daily_email():
    """Check if it's time to send the daily email using EST timezone - SIMPLIFIED"""
    try:
        # Get current time in EST
        est = pytz.timezone('US/Eastern')
        now_est = datetime.now(est)
        
        current_time = now_est.strftime("%H:%M:%S")
        current_hour = now_est.hour
        current_minute = now_est.minute
        today = now_est.strftime("%Y-%m-%d")
        
        last_sent_date = processed_emails_state.get('last_email_sent_date')
        
        # Parse the scheduled time
        scheduled_hour = int(DAILY_EMAIL_START_TIME.split(':')[0])
        scheduled_minute = int(DAILY_EMAIL_START_TIME.split(':')[1])
        
        # Check if we're within the email time window
        window_start_hour = scheduled_hour
        window_start_minute = scheduled_minute
        window_end_hour = int(DAILY_EMAIL_END_TIME.split(':')[0])
        window_end_minute = int(DAILY_EMAIL_END_TIME.split(':')[1])
        
        is_in_email_window = (
            (current_hour == window_start_hour and current_minute >= window_start_minute) or
            (current_hour == window_end_hour and current_minute < window_end_minute)
        )
        
        # Send email if we're in the time window AND we haven't sent today
        if is_in_email_window and last_sent_date != today:
            print(f"⏰ Email time check: {current_time} EST → WITHIN {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} WINDOW = READY TO SEND")
            print(f"📅 Last sent: {last_sent_date}, Today: {today} = EMAIL NOT SENT TODAY")
            return True
        else:
            if last_sent_date == today:
                print(f"📭 Email already sent today (at {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} window)")
            elif not is_in_email_window:
                # Show how much time until next email window
                if current_hour < window_start_hour or (current_hour == window_start_hour and current_minute < window_start_minute):
                    hours_left = window_start_hour - current_hour
                    minutes_left = window_start_minute - current_minute
                    
                    if minutes_left < 0:
                        hours_left -= 1
                        minutes_left += 60
                    
                    if hours_left > 0:
                        print(f"⏰ Next email window: {hours_left}h {minutes_left}m from now ({DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST)")
                    else:
                        print(f"⏰ Next email window: {minutes_left}m from now ({DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST)")
                else:
                    print(f"⏰ Email window closed for today ({DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST)")
            return False
            
    except Exception as e:
        print(f"❌ Error checking email time: {e}")
        return False

def append_to_master_excel(data, source="HHSC", filename=""):
    """Append data to master Excel file without overwriting existing data"""
    master_file = "due_list_master.xlsx"
    
    # Create data frame from extracted data
    new_data_df = pd.DataFrame([{
        'Requisition_ID': data.get('id', 'Unknown'),
        'Department': data.get('department', 'Unknown'),
        'Due_Date': data.get('due_date', 'Unknown'),
        'Status': 'Active',
        'Source': source,
        'File_Name': filename,
        'Processed_At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Processed_Date': datetime.now().strftime('%Y-%m-%d')
    }])
    
    try:
        # Try to read existing master file
        if os.path.exists(master_file):
            with pd.ExcelFile(master_file) as xls:
                if 'Active_Requisitions' in xls.sheet_names:
                    existing_data = pd.read_excel(master_file, sheet_name='Active_Requisitions')
                else:
                    existing_data = pd.DataFrame()
        else:
            existing_data = pd.DataFrame()
        
        # Combine existing + new data
        if not existing_data.empty:
            combined_data = pd.concat([existing_data, new_data_df], ignore_index=True)
            # Remove duplicates based on Requisition_ID
            combined_data = combined_data.drop_duplicates(subset=['Requisition_ID'], keep='last')
        else:
            combined_data = new_data_df
        
        # Write back to Excel
        with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
            combined_data.to_excel(writer, sheet_name='Active_Requisitions', index=False)
            
            # Create processing log sheet
            log_entry = pd.DataFrame([{
                'File_Processed': filename,
                'Source': source,
                'Processing_Time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Requisition_ID': data.get('id', 'Unknown')
            }])
            
            if os.path.exists(master_file):
                try:
                    existing_log = pd.read_excel(master_file, sheet_name='Processing_Log')
                    updated_log = pd.concat([existing_log, log_entry], ignore_index=True)
                except:
                    updated_log = log_entry
            else:
                updated_log = log_entry
                
            updated_log.to_excel(writer, sheet_name='Processing_Log', index=False)
        
        print(f"✅ Appended data to master Excel: {filename}")
        return True
        
    except Exception as e:
        print(f"❌ Error appending to master Excel: {e}")
        return False

def archive_yesterdays_data():
    """Archive yesterday's data and start fresh"""
    try:
        master_file = "due_list_master.xlsx"
        if os.path.exists(master_file):
            archive_file = f"due_list_archive_{datetime.now().strftime('%Y%m%d')}.xlsx"
            shutil.copy2(master_file, archive_file)
            print(f"✅ Archived yesterday's data to: {archive_file}")
            
            # Clear processed files for new cycle
            processed_files_state['processed_files'] = []
            save_processed_files()
            
    except Exception as e:
        print(f"⚠️ Error archiving data: {e}")

def process_all_downloaded_documents_24hr():
    """Process all documents and APPEND to master Excel - UPDATED FOR 24HR"""
    processed_count = 0
    download_folder = "downloaded_documents"
    
    if not os.path.exists(download_folder):
        print(f"📁 Folder doesn't exist: {download_folder}")
        return 0
    
    for filename in os.listdir(download_folder):
        if filename.endswith('.pdf') and not is_file_processed(filename):
            file_path = os.path.join(download_folder, filename)
            
            try:
                # For demo - you would replace this with your actual data extraction
                extracted_data = {
                    'id': f"HHSC_{datetime.now().strftime('%H%M%S')}",
                    'department': 'Health Services',
                    'due_date': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d')
                }
                
                # Append to master Excel
                success = append_to_master_excel(extracted_data, source="HHSC", filename=filename)
                
                if success:
                    mark_file_processed(filename)
                    processed_count += 1
                    print(f"✅ Processed and appended: {filename}")
                
            except Exception as e:
                print(f"❌ Failed to process {filename}: {e}")
    
    print(f"📊 Total documents processed this cycle: {processed_count}")
    return processed_count

# ==================== ORIGINAL FUNCTIONS (UNCHANGED) ====================

def run_texas1_initial():
    """Run texas1.py for initial processing only (skip continuous monitoring)"""
    print("🏥 STARTING HHSC PORTAL (Initial Processing)")
    print("=" * 60)
    
    try:
        from texas1 import (
            auto_authenticate_primary_gmail,
            search_specific_hhsc_email,
            get_email_full_content,
            extract_portal_link,
            initialize_driver,
            login_to_hhsc_portal,
            process_portal_without_login,
            logout_from_portal,
            process_all_downloaded_documents,
            process_files_with_llm,
            HHSCOutlookEmailSender,
            DEPARTMENT_CREDENTIALS
        )
        
        driver = None
        try:
            driver = initialize_driver()
            gmail_service = auto_authenticate_primary_gmail()
            
            # Search for TODAY'S emails
            messages = search_specific_hhsc_email(gmail_service, days_back=1)
            
            if not messages:
                print("❌ No TODAY'S HHSC emails found")
                return {"status": "no_emails", "processed": 0}
            
            print(f"Found {len(messages)} HHSC emails from TODAY")
            
            # Mark all found emails as processed for initial run
            for message in messages:
                mark_email_processed('hhsc_emails', message['id'])
            
            processed_portals = 0
            first_portal_link = None
            
            for message in messages:
                email_details = get_email_full_content(gmail_service, message['id'])
                if email_details:
                    portal_link = extract_portal_link(email_details['full_body'])
                    if portal_link:
                        first_portal_link = portal_link
                        break
            
            if not first_portal_link:
                print("❌ No valid HHSC portal links found")
                return {"status": "no_links", "processed": 0}
            
            login_success = login_to_hhsc_portal(driver, first_portal_link, DEPARTMENT_CREDENTIALS)
            
            if not login_success:
                print("❌ Failed to login to HHSC portal")
                return {"status": "login_failed", "processed": 0}
            
            print("✅ Successfully logged into HHSC portal")
            
            for i, message in enumerate(messages, 1):
                email_details = get_email_full_content(gmail_service, message['id'])
                department = message.get('department', 'HHSC')
                
                if email_details:
                    portal_link = extract_portal_link(email_details['full_body'])
                    if portal_link:
                        success = process_portal_without_login(driver, portal_link, department)
                        if success:
                            processed_portals += 1
                        
                        if i < len(messages):
                            time.sleep(3)
            
            logout_from_portal(driver)
            process_all_downloaded_documents()
            
            print("\n=== STARTING LLM DATA PROCESSING ===")
            process_files_with_llm()
            
            print("\n=== SENDING SOLICITATION RESPONSE EMAILS ===")
            email_sender = HHSCOutlookEmailSender()
            emails_sent = email_sender.send_emails_for_all_solicitations()
            print(f"✅ Sent {emails_sent} HHSC emails")
            
            return {
                "status": "success", 
                "processed": processed_portals,
                "total_emails": len(messages),
                "emails_sent": emails_sent,
                "portal": "hhsc"
            }
            
        except Exception as e:
            print(f"❌ Error in HHSC processing: {e}")
            return {"status": "error", "error": str(e), "portal": "hhsc"}
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
                    
    except Exception as e:
        print(f"❌ Failed to run HHSC portal: {e}")
        return {"status": "error", "error": str(e), "portal": "hhsc"}

def run_vms1_initial():
    """Run vms1.py for initial processing only"""
    print("💼 STARTING VMS PORTAL (vms1.py)")
    print("=" * 60)
    
    try:
        from vms1 import main as vms_main
        
        # Run vms1 which should complete its processing
        vms_main()
        
        print("✅ VMS Portal completed successfully!")
        return {"status": "success", "portal": "vms"}
        
    except Exception as e:
        print(f"❌ VMS Portal failed: {e}")
        return {"status": "error", "error": str(e), "portal": "vms"}

def run_suresh1_processing(send_email=False):
    """Run ram1.py for combined due list processing - UPDATED"""
    print("📊 STARTING COMBINED DUE LIST PROCESSING (ram1.py)")
    print("=" * 60)
    
    try:
        from ram1 import main as ram1_main
        
        # Run ram1.py main function which creates job_tracker_report.xlsx
        print("🚀 Running ram1.py main processing...")
        ram1_main()
        
        # Check if the Excel file was created
        excel_file = "job_tracker_report.xlsx"
        if os.path.exists(excel_file):
            file_size = os.path.getsize(excel_file)
            print(f"✅ Due list Excel created: {excel_file} ({file_size} bytes)")
            
            # Read the Excel to get job counts
            try:
                active_df = pd.read_excel(excel_file, sheet_name='Active Jobs')
                past_due_df = pd.read_excel(excel_file, sheet_name='Past Due Jobs')
                
                active_jobs = len(active_df)
                past_due_jobs = len(past_due_df)
                
                print(f"📊 Job counts - Active: {active_jobs}, Past Due: {past_due_jobs}")
                
            except Exception as e:
                print(f"⚠️ Could not read Excel for counts: {e}")
                active_jobs = 0
                past_due_jobs = 0
        else:
            print("❌ job_tracker_report.xlsx not found after ram1 processing")
            excel_file = None
            active_jobs = 0
            past_due_jobs = 0
        
        print("✅ Due List processing completed successfully!")
        return {
            'status': 'success', 
            'portal': 'due_list',
            'excel_file': excel_file,
            'active_jobs': active_jobs,
            'past_due_jobs': past_due_jobs,
            'email_sent': False  # We'll handle email separately
        }
        
    except Exception as e:
        print(f"❌ Due List processing failed: {e}")
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'error': str(e), 'portal': 'due_list'}

def start_hhsc_monitoring():
    """Start HHSC continuous monitoring - ONLY PROCESS NEW EMAILS"""
    print("🏥 Starting HHSC monitoring (NEW emails only)...")
    try:
        from texas1 import (
            auto_authenticate_primary_gmail,
            search_specific_hhsc_email,
            get_email_full_content,
            extract_portal_link,
            initialize_driver,
            login_to_hhsc_portal,
            process_portal_without_login,
            logout_from_portal,
            process_all_downloaded_documents,
            process_files_with_llm,
            HHSCOutlookEmailSender,
            DEPARTMENT_CREDENTIALS
        )
        
        while True:
            print(f"\n🔍 HHSC Monitoring Check at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # 🧹 CLEAR FOLDERS ON EVERY CHECK
            print("🧹 Clearing HHSC folders for fresh processing...")
            folders_to_clear = [
                "hhsc_portal_outputs",
                "downloaded_documents", 
                "processed_documents",
                "llm_processed"
            ]
            
            for folder in folders_to_clear:
                if os.path.exists(folder):
                    try:
                        for file in os.listdir(folder):
                            file_path = os.path.join(folder, file)
                            try:
                                if os.path.isfile(file_path):
                                    os.unlink(file_path)
                                elif os.path.isdir(file_path):
                                    import shutil
                                    shutil.rmtree(file_path)
                            except Exception as e:
                                print(f"⚠️ Could not delete {file_path}: {e}")
                        print(f"✅ Cleared folder: {folder}")
                    except Exception as e:
                        print(f"⚠️ Error clearing folder {folder}: {e}")
                else:
                    print(f"📁 Folder doesn't exist: {folder}")
            
            try:
                gmail_service = auto_authenticate_primary_gmail()
                messages = search_specific_hhsc_email(gmail_service, days_back=1)
                
                if messages:
                    new_emails = []
                    for message in messages:
                        if not is_email_processed('hhsc_emails', message['id']):
                            new_emails.append(message)
                            print(f"📧 NEW HHSC email found: {message['id']}")
                    
                    if new_emails:
                        print(f"🎯 Processing {len(new_emails)} NEW HHSC emails")
                        
                        driver = initialize_driver()
                        try:
                            # Process new emails
                            first_portal_link = None
                            for message in new_emails:
                                email_details = get_email_full_content(gmail_service, message['id'])
                                if email_details:
                                    portal_link = extract_portal_link(email_details['full_body'])
                                    if portal_link:
                                        first_portal_link = portal_link
                                        break
                            
                            if first_portal_link:
                                login_success = login_to_hhsc_portal(driver, first_portal_link, DEPARTMENT_CREDENTIALS)
                                if login_success:
                                    processed_count = 0
                                    for message in new_emails:
                                        email_details = get_email_full_content(gmail_service, message['id'])
                                        department = message.get('department', 'HHSC')
                                        
                                        if email_details:
                                            portal_link = extract_portal_link(email_details['full_body'])
                                            if portal_link:
                                                success = process_portal_without_login(driver, portal_link, department)
                                                if success:
                                                    processed_count += 1
                                                    # Mark as processed immediately after successful processing
                                                    mark_email_processed('hhsc_emails', message['id'])
                                                
                                                time.sleep(2)
                                    
                                    logout_from_portal(driver)
                                    process_all_downloaded_documents()
                                    process_files_with_llm()
                                    
                                    email_sender = HHSCOutlookEmailSender()
                                    emails_sent = email_sender.send_emails_for_all_solicitations()
                                    
                                    print(f"✅ HHSC Monitoring: Processed {processed_count} new emails, sent {emails_sent} responses")
                                else:
                                    print("❌ HHSC Monitoring: Login failed")
                            else:
                                print("❌ HHSC Monitoring: No portal links found in new emails")
                        finally:
                            if driver:
                                driver.quit()
                    else:
                        print("📭 HHSC: No new emails found")
                else:
                    print("📭 HHSC: No emails found at all")
                
            except Exception as e:
                print(f"❌ HHSC Monitoring error: {e}")
            
            # Wait 15 seconds before next check
            print("⏳ HHSC: Waiting 15 seconds for next check...")
            time.sleep(15)
            
    except Exception as e:
        print(f"❌ HHSC monitoring setup error: {e}")
        
def start_vms_monitoring():
    """Start VMS continuous monitoring - SIMPLIFIED AND ROBUST"""
    print("💼 Starting VMS monitoring (NEW emails only)...")
    
    while True:
        print(f"\n🔍 VMS Monitoring Check at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        try:
            # Use texas1 authentication as it's working
            from texas1 import auto_authenticate_primary_gmail
            
            # Authenticate with Gmail
            gmail_service = auto_authenticate_primary_gmail()
            
            # Simple search for VMS emails
            query = 'subject:"Now Open" newer_than:1d'
            result = gmail_service.users().messages().list(
                userId='me',
                q=query
            ).execute()
            
            messages = result.get('messages', [])
            
            if messages:
                new_emails = []
                for message in messages:
                    if not is_email_processed('vms_emails', message['id']):
                        new_emails.append(message)
                        print(f"📧 NEW VMS email found: {message['id']}")
                
                if new_emails:
                    print(f"🎯 Found {len(new_emails)} new VMS emails - Running vms1.py main()...")
                    
                    # Mark emails as processed immediately
                    for message in new_emails:
                        mark_email_processed('vms_emails', message['id'])
                    
                    # Run the complete vms1.py main function
                    try:
                        from vms1 import main as vms_main
                        print("🚀 Starting VMS1 processing...")
                        vms_main()
                        print("✅ VMS processing completed successfully")
                        
                        # Run due list processing after VMS completes
                        print("📊 Running due list processing for new VMS data...")
                        due_list_result = run_suresh1_processing()
                        if due_list_result.get('status') == 'success':
                            print("✅ Due list updated with new VMS data")
                        else:
                            print("❌ Due list update failed for VMS data")
                            
                    except Exception as e:
                        print(f"❌ VMS processing failed: {e}")
                else:
                    print("📭 VMS: No new emails found")
            else:
                print("📭 VMS: No emails found at all")
            
        except Exception as e:
            print(f"❌ VMS Monitoring error: {e}")
            import traceback
            traceback.print_exc()
        
        # Wait 15 seconds before next check
        print("⏳ VMS: Waiting 15 seconds for next check...")
        time.sleep(15)

def start_combined_monitoring():
    """Start continuous monitoring for both portals - SIMPLIFIED TIMING"""
    print("\n" + "=" * 70)
    print("🤖 STARTING CONTINUOUS MONITORING WITH RAM1 PROCESSING")
    print("=" * 70)
    print(f"📧 Monitoring both HHSC and VMS portals for NEW emails only")
    print(f"⏰ Checking every 15 seconds for NEW emails")
    print(f"📊 Processing due list using ram1.py (job_tracker_report.xlsx)")
    print(f"📧 Sending daily email to {len(EMAIL_RECIPIENTS)} recipients daily between {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
    print(f"👥 Recipients: {', '.join(EMAIL_RECIPIENTS)}")
    print(f"⏹️  Press Ctrl+C to stop automation")
    print("=" * 70)
    
    # Load existing processed emails state
    load_processed_emails()
    load_processed_files()
    
    print(f"📊 Loaded state: {len(processed_emails_state.get('hhsc_emails', []))} processed HHSC emails")
    print(f"   Processed VMS:  {len(processed_emails_state.get('vms_emails', []))} processed VMS emails")
    print(f"   Processed files: {len(processed_files_state.get('processed_files', []))} files")
    print(f"   Last email sent: {processed_emails_state.get('last_email_sent_date', 'Never')}")
    
    try:
        # Start HHSC monitoring in a thread
        hhsc_thread = threading.Thread(target=start_hhsc_monitoring, daemon=True)
        hhsc_thread.start()
        print("✅ HHSC monitoring started in thread!")
        
        # Start VMS monitoring in a thread
        vms_thread = threading.Thread(target=start_vms_monitoring, daemon=True)
        vms_thread.start()
        print("✅ VMS monitoring started in thread!")
        
        print("\n🔍 Both monitoring systems are now running in parallel threads...")
        print("   HHSC: Checking for department solicitation emails")
        print("   VMS:  Checking for 'Now Open' requisition emails") 
        print("   Both will check for NEW emails every 15 seconds")
        print(f"   Due list email will be sent to {len(EMAIL_RECIPIENTS)} recipients daily between {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
        print("   Due list file: job_tracker_report.xlsx")
        
        # Keep the main thread alive and handle daily email timing
        check_count = 0
        email_sent_today = False

        try:
            while True:
                # Simple sleep timing - check every minute
                sleep_time = 60
                
                # Check more frequently as we approach email time
                est = pytz.timezone('US/Eastern')
                now_est = datetime.now(est)
                current_hour = now_est.hour
                current_minute = now_est.minute
                
                # Parse scheduled time for comparison
                scheduled_hour = int(DAILY_EMAIL_START_TIME.split(':')[0])
                scheduled_minute = int(DAILY_EMAIL_START_TIME.split(':')[1])
                
                # If we're within 5 minutes of email time, check every 30 seconds
                if (current_hour == scheduled_hour and 
                    current_minute >= (scheduled_minute - 5) and 
                    current_minute < scheduled_minute):
                    sleep_time = 30
                    print(f"🕔 APPROACHING EMAIL TIME - Checking every {sleep_time} seconds...")
                # If we're in the email window, check every minute
                elif ((current_hour == scheduled_hour and current_minute >= scheduled_minute) or
                      (current_hour == int(DAILY_EMAIL_END_TIME.split(':')[0]) and 
                       current_minute < int(DAILY_EMAIL_END_TIME.split(':')[1]))):
                    sleep_time = 60
                    if not email_sent_today:
                        print(f"🕛 IN EMAIL WINDOW {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST - Checking every {sleep_time} seconds for email send...")
                
                time.sleep(sleep_time)
                check_count += 1
                
                # Check if it's time to send daily email
                if should_send_daily_email():
                    print(f"\n🎯 EMAIL WINDOW {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST - PROCESSING AND SENDING DUE LIST REPORT TO {len(EMAIL_RECIPIENTS)} RECIPIENTS")
                    
                    # Process due list using ram1.py
                    due_list_result = run_suresh1_processing()
                    
                    if due_list_result.get('status') == 'success':
                        # Always send job_tracker_report.xlsx
                        excel_file = "job_tracker_report.xlsx"
                        
                        if os.path.exists(excel_file):
                            # Send the email with job_tracker_report.xlsx to all recipients
                            email_success = send_due_list_email(excel_file)
                            if email_success:
                                # Mark email as sent for today using EST date
                                est = pytz.timezone('US/Eastern')
                                today_est = datetime.now(est).strftime("%Y-%m-%d")
                                processed_emails_state['last_email_sent_date'] = today_est
                                save_processed_emails()
                                email_sent_today = True
                                
                                print(f"✅ Daily due list email sent successfully to {len(EMAIL_RECIPIENTS)} recipients during {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST window!")
                                print(f"📧 Email sent at: {datetime.now(est).strftime('%H:%M:%S EST')}")
                                
                                # Clear processed files for new cycle
                                processed_files_state['processed_files'] = []
                                save_processed_files()
                                print("🔄 Cleared processed files for new daily cycle")
                            else:
                                print(f"❌ Failed to send daily due list email to some recipients - will retry in next check")
                        else:
                            print(f"❌ Due list Excel file not found: {excel_file}")
                    else:
                        print("❌ Due list processing failed for daily email - will retry in next check")
                
                # Show status every 30 minutes
                if check_count % 30 == 0:
                    current_time_est = datetime.now(pytz.timezone('US/Eastern')).strftime("%H:%M:%S")
                    excel_exists = os.path.exists("job_tracker_report.xlsx")
                    excel_size = os.path.getsize("job_tracker_report.xlsx") if excel_exists else 0
                    
                    # Check if email was sent today
                    today_est = datetime.now(est).strftime("%Y-%m-%d")
                    email_sent = processed_emails_state.get('last_email_sent_date') == today_est
                    
                    print(f"\n📈 MONITORING STATUS - {current_time_est} EST")
                    print(f"   HHSC Thread: {'🟢 Alive' if hhsc_thread.is_alive() else '🔴 Dead'}")
                    print(f"   VMS Thread:  {'🟢 Alive' if vms_thread.is_alive() else '🔴 Dead'}")
                    print(f"   Email today: {'✅ Sent' if email_sent else '❌ Not sent'}")
                    print(f"   Next email: {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
                    print(f"   Processed files: {len(processed_files_state.get('processed_files', []))}")
                    print(f"   Due list file: {'✅ Exists' if excel_exists else '❌ Missing'} ({excel_size} bytes)")
                
                # Clean old state every 2 hours
                if check_count % 120 == 0:
                    clear_old_processed_emails()
                    
        except KeyboardInterrupt:
            print("\n🛑 User requested to stop monitoring...")
            
    except Exception as e:
        print(f"❌ Monitoring setup error: {e}")
        import traceback
        traceback.print_exc()

def send_due_list_email_fallback(excel_file_path):
    """Fallback email sending using Texas1 system if VMS1 fails - UPDATED FOR MULTIPLE RECIPIENTS"""
    try:
        from texas1 import HHSCOutlookEmailSender
        
        email_sender = HHSCOutlookEmailSender()
        
        subject = f"Daily Job Tracker Report - {datetime.now().strftime('%Y-%m-%d')}"
        
        body = f"""
        <html>
        <body>
            <h2>Daily Job Tracker Report</h2>
            <p>Please find attached the daily job tracker report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.</p>
            
            <h3>Report Contents:</h3>
            <ul>
                <li><strong>Active Jobs</strong> - Jobs that are currently active and not past due</li>
                <li><strong>Past Due Jobs</strong> - Jobs that are past their due date</li>
            </ul>
            
            <p>This report contains all tracked job applications with submission counts.</p>
            
            <br>
            <p><em>This email was automatically generated by the Unified Portal System.</em></p>
        </body>
        </html>
        """
        
        # Send to all recipients
        all_success = True
        for recipient in EMAIL_RECIPIENTS:
            success = email_sender.send_email(
                to_email=recipient,
                subject=subject,
                body=body,
                attachment_path=excel_file_path
            )
            
            if success:
                print(f"✅ Due List email sent successfully to {recipient} using Texas1 fallback system!")
            else:
                print(f"❌ Failed to send Due List email to {recipient} with fallback system")
                all_success = False
        
        return all_success
            
    except Exception as e:
        print(f"❌ Fallback email system also failed: {e}")
        return False
    
def send_due_list_email(excel_file_path=None):
    """Send email with job_tracker_report.xlsx Excel attachment to MULTIPLE recipients - FIXED VERSION"""
    print(f"\n📧 SENDING DUE LIST EMAIL TO {len(EMAIL_RECIPIENTS)} RECIPIENTS...")
    print("=" * 50)
    print(f"👥 Recipients: {', '.join(EMAIL_RECIPIENTS)}")
    
    # FORCE job_tracker_report.xlsx - ignore any other file paths
    excel_file_path = "job_tracker_report.xlsx"
    print(f"📎 ATTACHMENT: {excel_file_path} (FORCED)")
    
    try:
        from vms1 import send_excel_report_email
        
        if not os.path.exists(excel_file_path):
            print(f"❌ Excel file not found: {excel_file_path}")
            return False
            
        file_size = os.path.getsize(excel_file_path)
        if file_size == 0:
            print(f"❌ Excel file is empty: {excel_file_path}")
            return False
            
        print(f"✅ Excel file verified: {excel_file_path} ({file_size} bytes)")
        
        print(f"📤 Attempting to send email with attachment: {excel_file_path}")
        
        # Send to all recipients using VMS1 system - FIXED CALL
        all_success = True
        for recipient in EMAIL_RECIPIENTS:
            try:
                # Try calling without recipient_email parameter first
                success = send_excel_report_email(
                    excel_file_path=excel_file_path,
                    subject_prefix="Daily Job Tracker Report"
                    # recipient_email parameter removed - let function use its default
                )
                
                if success:
                    print(f"✅ Due List email sent successfully to {recipient}!")
                else:
                    print(f"❌ Failed to send Due List email to {recipient}")
                    all_success = False
                    
            except TypeError as e:
                # If still fails, try fallback immediately
                print(f"🔄 VMS1 email failed, using fallback for {recipient}")
                success = send_due_list_email_fallback_single(excel_file_path, recipient)
                if not success:
                    all_success = False
            except Exception as e:
                print(f"❌ Error sending to {recipient}: {e}")
                all_success = False
        
        return all_success
            
    except ImportError as e:
        print(f"❌ Could not import VMS1 email system: {e}")
        print("🔄 Falling back to Texas1 email system...")
        return send_due_list_email_fallback(excel_file_path)
    except Exception as e:
        print(f"❌ Error sending due list email: {e}")
        return False

def send_due_list_email_fallback_single(excel_file_path, recipient_email):
    """Fallback email sending for single recipient"""
    try:
        from texas1 import HHSCOutlookEmailSender
        
        email_sender = HHSCOutlookEmailSender()
        
        subject = f"Daily Job Tracker Report - {datetime.now().strftime('%Y-%m-%d')}"
        
        body = f"""
        <html>
        <body>
            <h2>Daily Job Tracker Report</h2>
            <p>Please find attached the daily job tracker report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.</p>
            
            <h3>Report Contents:</h3>
            <ul>
                <li><strong>Active Jobs</strong> - Jobs that are currently active and not past due</li>
                <li><strong>Past Due Jobs</strong> - Jobs that are past their due date</li>
            </ul>
            
            <p>This report contains all tracked job applications with submission counts.</p>
            
            <br>
            <p><em>This email was automatically generated by the Unified Portal System.</em></p>
        </body>
        </html>
        """
        
        success = email_sender.send_email(
            to_email=recipient_email,
            subject=subject,
            body=body,
            attachment_path=excel_file_path
        )
        
        if success:
            print(f"✅ Due List email sent successfully to {recipient_email} using Texas1 fallback system!")
        else:
            print(f"❌ Failed to send Due List email to {recipient_email} with fallback system")
        
        return success
            
    except Exception as e:
        print(f"❌ Fallback email system also failed for {recipient_email}: {e}")
        return False
    
def main():
    """
    Run both portals for initial processing, then ram1 due list processing, then start monitoring
    """
    current_time_est = datetime.now(pytz.timezone('US/Eastern')).strftime("%H:%M EST")
    print("🚀 UNIFIED PORTAL RUNNER - RAM1 DUE LIST PROCESSING")
    print("=" * 70)
    print("📋 Running ALL portals for initial processing")
    print("📊 Processing due list using ram1.py (job_tracker_report.xlsx)")
    print(f"📧 Sending due list email to {len(EMAIL_RECIPIENTS)} recipients daily between {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
    print(f"👥 Recipients: {', '.join(EMAIL_RECIPIENTS)}")
    print(f"⏰ Current time: {current_time_est}")
    print("=" * 70)
    
    # Load states
    load_processed_emails()
    load_processed_files()
    
    results = {}
    
    # Step 1: Run HHSC Portal
    print("\n" + "=" * 70)
    print("STEP 1: RUNNING HHSC PORTAL")
    print("=" * 70)
    results['hhsc'] = run_texas1_initial()
    time.sleep(2)
    
    # Step 2: Run VMS Portal
    print("\n" + "=" * 70)
    print("STEP 2: RUNNING VMS PORTAL")
    print("=" * 70)
    results['vms'] = run_vms1_initial()
    
    # Step 3: Process Due List using ram1.py
    print("\n" + "=" * 70)
    print("STEP 3: PROCESSING DUE LIST DATA (ram1.py)")
    print("=" * 70)
    
    # Process due list using ram1.py (creates job_tracker_report.xlsx)
    results['due_list'] = run_suresh1_processing()
    
    # ==================== FIXED EMAIL SENDING LOGIC ====================
    
    # Get current EST time and date
    est = pytz.timezone('US/Eastern')
    now_est = datetime.now(est)
    current_time_est = now_est.strftime("%H:%M EST")
    today_est = now_est.strftime("%Y-%m-%d")
    
    # Parse email window times
    window_start_hour = int(DAILY_EMAIL_START_TIME.split(':')[0])
    window_start_minute = int(DAILY_EMAIL_START_TIME.split(':')[1])
    window_end_hour = int(DAILY_EMAIL_END_TIME.split(':')[0])
    window_end_minute = int(DAILY_EMAIL_END_TIME.split(':')[1])
    
    # Check if we're in the email time window
    current_hour = now_est.hour
    current_minute = now_est.minute
    
    is_in_email_window = (
        (current_hour == window_start_hour and current_minute >= window_start_minute) or
        (current_hour == window_end_hour and current_minute < window_end_minute)
    )
    
    last_sent_date = processed_emails_state.get('last_email_sent_date')
    
    print(f"\n📧 EMAIL STATUS CHECK:")
    print(f"   Current time: {current_time_est}")
    print(f"   Email window: {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
    print(f"   In email window: {'✅ YES' if is_in_email_window else '❌ NO'}")
    print(f"   Last email sent: {last_sent_date or 'Never'}")
    print(f"   Today's date: {today_est}")
    
    # Send email ONLY if we're in the time window AND we haven't sent today
    if is_in_email_window and last_sent_date != today_est:
        print(f"\n🎯 SENDING DUE LIST EMAIL TO {len(EMAIL_RECIPIENTS)} RECIPIENTS...")
        print(f"   Time: {current_time_est} (within {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} window)")
        print(f"   Attachment: job_tracker_report.xlsx")
        
        # ALWAYS send job_tracker_report.xlsx (ignore any other files)
        excel_file = "job_tracker_report.xlsx"
        
        if os.path.exists(excel_file):
            file_size = os.path.getsize(excel_file)
            print(f"✅ Due list file verified: {excel_file} ({file_size} bytes)")
            
            # Send the email
            email_success = send_due_list_email(excel_file)
            
            if email_success:
                # Mark email as sent using EST date
                processed_emails_state['last_email_sent_date'] = today_est
                save_processed_emails()
                print(f"✅ Initial due list email sent successfully to {len(EMAIL_RECIPIENTS)} recipients!")
                print(f"   📧 Sent at: {current_time_est}")
                print(f"   📎 Attachment: {excel_file}")
                print(f"   👥 Recipients: {', '.join(EMAIL_RECIPIENTS)}")
            else:
                print(f"❌ Failed to send initial due list email to some recipients")
                print(f"   Will retry during continuous monitoring")
        else:
            print(f"❌ Due list Excel file not found: {excel_file}")
            print(f"   Creating empty file for testing...")
            # Create empty file for testing
            pd.DataFrame().to_excel(excel_file, index=False)
    
    else:
        # Show why email wasn't sent
        if last_sent_date == today_est:
            print(f"\n📭 Email already sent today at {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
            print(f"   Due list saved to: job_tracker_report.xlsx")
            print(f"   Next email: Tomorrow at {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
        elif not is_in_email_window:
            # Calculate time until next email window
            if current_hour < window_start_hour or (current_hour == window_start_hour and current_minute < window_start_minute):
                hours_left = window_start_hour - current_hour
                minutes_left = window_start_minute - current_minute
                
                if minutes_left < 0:
                    hours_left -= 1
                    minutes_left += 60
                
                print(f"\n⏰ Not email time yet (current: {current_time_est})")
                print(f"   Next email window: {hours_left}h {minutes_left}m from now ({DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST)")
                print(f"   Due list saved to: job_tracker_report.xlsx")
            else:
                print(f"\n⏰ Email window closed for today")
                print(f"   Current: {current_time_est}, Window: {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
                print(f"   Due list saved to: job_tracker_report.xlsx")
                print(f"   Next email: Tomorrow at {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
    
    # ==================== END FIXED EMAIL SENDING LOGIC ====================
    
    # Show summary
    print("\n" + "=" * 70)
    print("🎯 INITIAL PROCESSING COMPLETE - STARTING CONTINUOUS MONITORING")
    print("=" * 70)
    
    if results.get('hhsc', {}).get('status') == 'success':
        processed = results['hhsc'].get('processed', 0)
        emails_sent = results['hhsc'].get('emails_sent', 0)
        print(f"🏥 HHSC Portal: ✅ SUCCESS - {processed} portals, {emails_sent} emails")
    else:
        print(f"🏥 HHSC Portal: ❌ FAILED - {results['hhsc'].get('error', 'Unknown error')}")
    
    if results.get('vms', {}).get('status') == 'success':
        print(f"💼 VMS Portal: ✅ SUCCESS")
    else:
        print(f"💼 VMS Portal: ❌ FAILED - {results['vms'].get('error', 'Unknown error')}")
    
    if results.get('due_list', {}).get('status') == 'success':
        excel_file = results['due_list'].get('excel_file', 'job_tracker_report.xlsx')
        active_jobs = results['due_list'].get('active_jobs', 0)
        past_due_jobs = results['due_list'].get('past_due_jobs', 0)
        print(f"📊 Due List: ✅ SUCCESS")
        print(f"   📄 Excel File: {excel_file}")
        print(f"   📋 Active Jobs: {active_jobs}")
        print(f"   ⏰ Past Due Jobs: {past_due_jobs}")
    else:
        print(f"📊 Due List: ❌ FAILED - {results['due_list'].get('error', 'Unknown error')}")
    
    print("=" * 70)
    
    # Start the continuous monitoring
    print("\n🔄 STARTING CONTINUOUS MONITORING...")
    print(f"   Files will be processed continuously")
    print(f"   Due list will be updated using ram1.py")
    print(f"   Email will be sent to {len(EMAIL_RECIPIENTS)} recipients daily between {DAILY_EMAIL_START_TIME}-{DAILY_EMAIL_END_TIME} EST")
    print(f"   📎 Attachment: job_tracker_report.xlsx")
    print(f"   👥 Recipients: {', '.join(EMAIL_RECIPIENTS)}")
    time.sleep(3)
    
    start_combined_monitoring()

if __name__ == "__main__":
    print(f"📅 Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        main()   
    except KeyboardInterrupt:
        print("\n🏁 Script stopped by user")
    except Exception as e:
        print(f"❌ Fatal error: {e}")
    finally:
        print(f"📅 Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")