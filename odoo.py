import xmlrpc.client
import os
import re
from datetime import datetime

# Odoo connection configuration
url = "http://84.247.136.24:8069"
db = 'mydb'
username = 'admin'
password = "79NM46eRqDv)w^q^"

def parse_due_date_from_job_id(job_id):
    """
    Extract and parse due date from Job ID format like (98091204)
    Returns formatted date string: "Monday, December 04, 2025"
    """
    try:
        print(f"   - Parsing Job ID for due date: {job_id}")
        
        # Extract the number in parentheses
        match = re.search(r'\((\d+)\)', job_id)
        if match:
            number_str = match.group(1)
            print(f"   - Extracted number: {number_str}")
            
            # Take last 4 digits
            if len(number_str) >= 4:
                date_code = number_str[-4:]  # Get last 4 digits
                print(f"   - Last 4 digits: {date_code}")
                
                # Parse month and day from MMDD format
                month = int(date_code[:2])
                day = int(date_code[2:])
                
                # Validate month and day
                if 1 <= month <= 12 and 1 <= day <= 31:
                    # Use current year or next year if date has passed
                    current_year = datetime.now().year
                    
                    # Try to create date
                    try:
                        due_date = datetime(current_year, month, day)
                        
                        # If date is in the past, use next year
                        if due_date < datetime.now():
                            due_date = datetime(current_year + 1, month, day)
                        
                        # Format as "Monday, December 04, 2025"
                        formatted_date = due_date.strftime("%A, %B %d, %Y")
                        print(f"   - Parsed date: {formatted_date}")
                        return formatted_date
                    except ValueError as ve:
                        # Invalid date (e.g., Feb 30)
                        print(f"   - Invalid date error: {ve}")
                        return ""
                else:
                    print(f"   - Invalid month/day: month={month}, day={day}")
            else:
                print(f"   - Number too short: {number_str}")
        else:
            print(f"   - No number found in parentheses")
        
        return ""
    except Exception as e:
        print(f"   - Error parsing due date: {e}")
        return ""

def get_or_create_contract_type(models, uid, password, employment_type):
    """
    Get existing contract type ID or create a new one based on employment type from .txt file
    """
    if not employment_type:
        return False
    
    try:
        # First, search if this contract type already exists
        contract_type_ids = models.execute_kw(
            db, uid, password,
            "hr.contract.type", "search",
            [[("name", "=", employment_type)]]
        )
        
        if contract_type_ids:
            print(f"   - Found existing contract type: '{employment_type}' (ID: {contract_type_ids[0]})")
            return contract_type_ids[0]
        else:
            # Create new contract type
            new_contract_type_id = models.execute_kw(
                db, uid, password,
                "hr.contract.type", "create",
                [{"name": employment_type}]
            )
            print(f"   - Created new contract type: '{employment_type}' (ID: {new_contract_type_id})")
            return new_contract_type_id
            
    except Exception as e:
        print(f"⚠️ Warning: Could not get/create contract type for '{employment_type}': {e}")
        return False

def extract_employment_type(content):
    """
    Extract employment type from content with proper priority order
    """
    # Define employment types in priority order (most specific first)
    employment_types = [
        "Hybrid/Local",
        "Onsite/Local", 
        "Remote/Local",
        "Hybrid",
        "Remote",
        "Onsite",
        "Local"
    ]
    
    # Look for each employment type in the content
    for emp_type in employment_types:
        if emp_type in content:
            return emp_type
    
    # Default if none found
    return "Hybrid/Local"

def extract_full_job_title(content):
    """
    Extract the full job title from the content
    """
    try:
        # Split content into lines
        lines = content.split('\n')
        
        # Look for the line that contains the full job title
        # This is typically the first line that starts with an employment type
        for line in lines:
            line = line.strip()
            # Check if line starts with any employment type pattern
            if (line.startswith("Hybrid/Local") or 
                line.startswith("Onsite/Local") or
                line.startswith("Remote/Local") or
                line.startswith("Hybrid") or
                line.startswith("Remote") or
                line.startswith("Onsite") or
                line.startswith("Local")):
                # Return the entire line as the full title
                return line
        
        # If no pattern found, try to find the first non-empty line that's not "Job ID:"
        for line in lines:
            if line.strip() and not line.strip().startswith("Job ID:"):
                return line.strip()
                
        return "Unknown Position"
        
    except Exception as e:
        print(f"   - Error extracting full job title: {e}")
        return "Unknown Position"

def extract_job_data_from_file(file_path):
    """
    Extract job posting data from text files
    """
    job_data = {}
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        print(f"   - File content preview: {content[:200]}...")  # Debug: show first 200 chars
            
        # Extract FULL Job Title (not just position)
        full_title = extract_full_job_title(content)
        job_data['name'] = full_title
        print(f"   - Full Job Title: {full_title}")
        
        # Extract Description
        desc_match = re.search(r'Description:\s*(.*?)(?=Skills:|$)', content, re.DOTALL | re.IGNORECASE)
        if desc_match:
            job_data['x_description'] = desc_match.group(1).strip()
        else:
            # If no specific description section, use the entire content
            job_data['x_description'] = content
        
        # Extract Target (Number of positions)
        positions_match = re.search(r'Positions:\s*(\d+)', content, re.IGNORECASE)
        if positions_match:
            job_data['no_of_recruitment'] = int(positions_match.group(1))
        else:
            job_data['no_of_recruitment'] = 1
        
        # Extract Job ID
        job_id_match = re.search(r'Job ID:\s*([^\n]+)', content)
        if job_id_match:
            job_id_full = job_id_match.group(1).strip()
            job_data['x_job_id'] = job_id_full
            
            # Extract and parse due date from Job ID
            due_date = parse_due_date_from_job_id(job_id_full)
            if due_date:
                job_data['x_due_date'] = due_date
            else:
                job_data['x_due_date'] = ""
        
        # Extract Location
        location_match = re.search(r'Location:\s*([^\n\(]+)', content)
        if location_match:
            job_data['x_location'] = location_match.group(1).strip()
        
        # Extract Duration
        duration_match = re.search(r'Duration:\s*([^\n]+)', content)
        if duration_match:
            job_data['x_duration'] = duration_match.group(1).strip()
        
        # Extract Skills
        skills_match = re.search(r'Skills:(.*?)(?=Description:|$)', content, re.DOTALL | re.IGNORECASE)
        if skills_match:
            job_data['x_skills'] = skills_match.group(1).strip()
        
        # Extract Employment Type
        job_data['employment_type'] = extract_employment_type(content)
        print(f"   - Detected employment type: {job_data['employment_type']}")
        
        # Extract Expected Skills (from the main title line)
        expected_skills_match = re.search(r'with\s+([^,]+(?:,[^,]+)*)', content)
        if expected_skills_match:
            job_data['expected_skills_text'] = expected_skills_match.group(1).strip()
            
    except Exception as e:
        print(f"   ❌ Error reading file {file_path}: {e}")
        return None
    
    return job_data

def check_existing_contract_types(models, uid, password):
    """
    Check what contract types exist in Odoo
    """
    try:
        contract_types = models.execute_kw(
            db, uid, password,
            "hr.contract.type", "search_read",
            [[]], {'fields': ['id', 'name']}
        )
        print("📋 Existing Contract Types in Odoo:")
        for ct in contract_types:
            print(f"   ID: {ct['id']}, Name: '{ct['name']}'")
        return contract_types
    except Exception as e:
        print(f"❌ Error fetching contract types: {e}")
        return []

def create_odoo_job_posting(job_data, models, uid, password):
    """
    Create job posting in Odoo with extracted data using exact field names
    """
    # Get or create contract type ID based on .txt file content
    employment_type = job_data.get('employment_type', '')
    contract_type_id = False
    
    if employment_type:
        contract_type_id = get_or_create_contract_type(models, uid, password, employment_type)

    # Prepare job data for Odoo - USING EXACT FIELD NAMES FROM INSPECT
    odoo_job_data = {
        # Standard Odoo fields - NOW WITH FULL TITLE
        "name": job_data.get('name', 'Unknown Position'),
        "no_of_recruitment": job_data.get('no_of_recruitment', 1),
        "company_id": 1,
        "department_id": False,
        "user_id": False,
        "website_published": True,
        
        # Custom fields from your inspect (x_ fields)
        "x_job_id": job_data.get('x_job_id', ''),
        "x_location": job_data.get('x_location', ''),
        "x_duration": job_data.get('x_duration', ''),
        "x_skills": job_data.get('x_skills', ''),
        "x_description": job_data.get('x_description', 'No description available.'),
        "x_due_date": job_data.get('x_due_date', ''),  # NEW: Due Date field (char field)
        
        # Contract type field
        "contract_type_id": contract_type_id,
        
        # Address field (set to company address)
        "address_id": 1,  # Default company address
    }
    
    # Debug: Show what data is being sent
    print(f"   - Data being sent to Odoo:")
    print(f"     x_due_date: '{odoo_job_data.get('x_due_date')}'")
    print(f"     x_due_date type: {type(odoo_job_data.get('x_due_date'))}")

    # Create job posting
    try:
        job_id = models.execute_kw(
            db, uid, password,
            "hr.job", "create", [odoo_job_data]
        )
        print(f"✅ Job posting created successfully. Job ID: {job_id}")
        print(f"   Full Title: {job_data.get('name')}")
        print(f"   Custom Job ID: {job_data.get('x_job_id')}")
        print(f"   Location: {job_data.get('x_location')}")
        print(f"   Due Date: '{job_data.get('x_due_date', 'Not specified')}'")
        print(f"   Employment Type: {job_data.get('employment_type', 'Not specified')}")
        print(f"   Contract Type ID: {contract_type_id}")
        return True
    except Exception as e:
        print(f"❌ Error creating job posting: {e}")
        return False

def process_job_files_from_folders():
    """
    Process all job posting files from specified folders
    """
    # Connect to Odoo server first (single connection for all files)
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
    uid = common.authenticate(db, username, password, {})

    if not uid:
        print("Authentication failed.")
        return

    # Connect to Odoo object endpoint
    models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
    
    # First, check what contract types exist
    print("🔍 Checking existing contract types...")
    existing_contract_types = check_existing_contract_types(models, uid, password)
    
    folders = ['requisition_outputs', 'hhsc_portal_outputs']
    processed_files = 0
    successful_postings = 0
    
    for folder in folders:
        if not os.path.exists(folder):
            print(f"⚠️ Folder '{folder}' does not exist. Skipping...")
            continue
            
        print(f"\n📁 Processing folder: {folder}")
        
        # Process all .txt files in the folder
        for filename in os.listdir(folder):
            if filename.endswith('.txt'):
                file_path = os.path.join(folder, filename)
                print(f"\n📄 Processing file: {filename}")
                
                # Extract job data from file
                job_data = extract_job_data_from_file(file_path)
                
                if job_data and job_data.get('name'):
                    print(f"   Extracted data:")
                    print(f"   - Full Title: {job_data.get('name')}")
                    print(f"   - Job ID: {job_data.get('x_job_id')}")
                    print(f"   - Due Date: '{job_data.get('x_due_date', 'Not extracted')}'")
                    print(f"   - Location: {job_data.get('x_location')}")
                    print(f"   - Positions: {job_data.get('no_of_recruitment')}")
                    print(f"   - Duration: {job_data.get('x_duration')}")
                    print(f"   - Employment Type: {job_data.get('employment_type', 'Not specified')}")
                    
                    # Create Odoo job posting
                    if create_odoo_job_posting(job_data, models, uid, password):
                        successful_postings += 1
                else:
                    print(f"   ❌ No valid job data extracted from {filename}")
                    # Show file content for debugging
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            print(f"   File content (first 500 chars): {content[:500]}...")
                    except Exception as e:
                        print(f"   Could not read file for debugging: {e}")
                
                processed_files += 1
    
    print(f"\n📊 Summary:")
    print(f"   Total files processed: {processed_files}")
    print(f"   Successful job postings: {successful_postings}")

# Main execution
if __name__ == "__main__":
    print("🚀 Starting Odoo Job Posting Automation")
    print("=" * 50)
    
    process_job_files_from_folders()
    
    print("\n🎯 Automation completed!")