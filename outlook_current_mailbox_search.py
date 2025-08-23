"""
Enhanced Outlook search script that searches the CURRENT MAILBOX instead of default inbox
Searches for:
1. reservations.gmhd@millenniumhotels.com
2. Avital Boaz emails
Shows results with all field details from the current mailbox
"""
import win32com.client
import pythoncom
import pandas as pd
from datetime import datetime, timedelta
import re
import pdfplumber
import io
import os

def connect_to_outlook():
    """Connect to Outlook using win32com.client"""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        print(f"Outlook connection failed: {e}")
        return None, None

def extract_pdf_text(pdf_bytes):
    """Extract text from PDF bytes"""
    try:
        pdf_file = io.BytesIO(pdf_bytes)
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except Exception as e:
        print(f"PDF extraction failed: {e}")
        return ""

def extract_reservation_fields(text, sender_email=""):
    """Extract reservation fields using regex patterns"""
    
    # Different patterns for different email sources
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        # Patterns for noreply-reservations emails based on actual structure
        patterns = {
            'GUEST_NAME_FULL': r"Guest Name:\s*(.+?)(?:\n|Address:)",
            'ARRIVAL': r"Arrive:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'DEPARTURE': r"Depart:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'NIGHTS': r"Total Nights\s*(\d+)\s*night",
            'PERSONS': r"Adult/Children:\s*(\d+)/\d+",
            'ROOM': r"Room with One King Bed\n(?:.*\n)?Room Code:\s*[A-Z0-9]+\s*(?:.*\n)?",
            'ROOM_TYPE': r"(Superior Room|Deluxe Room|Standard Room|Executive Room)",
            'RATE_CODE': r"Rate Code:\s*([A-Z0-9]+)",
            'RATE_NAME': r"Rate Name:\s*(.+?)(?:\n|Rate Code:)",
            'COMPANY': r"Travel Agent\s*(?:.*\n)*Name:\s*(.+?)(?:\n|$)",
            'NET_TOTAL': r"Total charges:\s*AED\s*([0-9,]+\.?[0-9]*)",
            'CONFIRMATION': r"Confirman:\s*([A-Z0-9]+)",
        }
        
        # Additional patterns to extract from subject line
        subject_patterns = {
            'ARRIVAL_SUBJECT': r"Arrival Date[:]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'CONFIRMATION_SUBJECT': r"confirmation number[:]*([A-Z0-9]+)",
        }
        
        # Merge subject patterns with main patterns
        patterns.update(subject_patterns)
    else:
        # Original patterns for other emails (PDFs, etc.)
        patterns = {
            'FULL_NAME': r"(?:Name|Guest Name)[:\s]+(.+?)(?:\n|$)",
            'FIRST_NAME': r"(?:First Name)[:\s]+(.+?)(?:\n|$)",
            'ARRIVAL': r"(?:Arrival|Check-in)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'DEPARTURE': r"(?:Departure|Check-out)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'NIGHTS': r"(?:Nights|Night)[:\s]+(\d+)",
            'PERSONS': r"(?:Persons|Guest|Adults?)[:\s]+(\d+)",
            'ROOM': r"(?:Room|Room Type)[:\s]+(.+?)(?:\n|$)",
            'RATE_CODE': r"(?:Rate Code|Rate)[:\s]+(.+?)(?:\n|$)",
            'COMPANY': r"(?:Company|Agency)[:\s]+(.+?)(?:\n|$)",
            'NET_TOTAL': r"(?:Total|Net Total|Amount|Net Amount)[:\s]+(?:AED\s*)?([\\d,]+\.?\\d*)",
        }
    
    extracted = {}
    
    # Extract all fields first
    for field, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            value = match.group(1).strip()
            extracted[field] = value
        else:
            extracted[field] = "N/A"
    
    # Special processing for noreply-reservations emails
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        # Process guest name - split "Boaz Avital" into first name and last name
        guest_name = extracted.get('GUEST_NAME_FULL', 'N/A')
        if guest_name != 'N/A' and guest_name.strip():
            name_parts = guest_name.strip().split()
            if len(name_parts) >= 2:
                # First name is the first part, full name (last name) is the last part
                extracted['FIRST_NAME'] = name_parts[0]
                extracted['FULL_NAME'] = name_parts[-1]  # Last name as full name per instruction
            else:
                extracted['FIRST_NAME'] = guest_name
                extracted['FULL_NAME'] = guest_name
        else:
            extracted['FIRST_NAME'] = 'N/A'
            extracted['FULL_NAME'] = 'N/A'
        
        # Map room types to codes
        room_type = extracted.get('ROOM_TYPE', 'N/A')
        if room_type != 'N/A':
            if 'Superior Room' in room_type:
                extracted['ROOM'] = 'SK'  # Map Superior Room to SK as requested
            else:
                extracted['ROOM'] = room_type
        else:
            extracted['ROOM'] = 'N/A'
        
        # Use rate code as primary, fallback to rate name
        if extracted.get('RATE_CODE', 'N/A') != 'N/A':
            # Keep the rate code as is
            pass
        elif extracted.get('RATE_NAME', 'N/A') != 'N/A':
            extracted['RATE_CODE'] = extracted['RATE_NAME']
        else:
            extracted['RATE_CODE'] = 'N/A'
    
    # Convert dates to dd/mm/yyyy format
    for date_field in ['ARRIVAL', 'DEPARTURE', 'ARRIVAL_SUBJECT']:
        if date_field in extracted and extracted[date_field] != 'N/A':
            try:
                parsed_date = pd.to_datetime(extracted[date_field], dayfirst=True)
                extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
            except:
                pass  # Keep original value if parsing fails
    
    # Use arrival from subject if main arrival not found
    if extracted.get('ARRIVAL', 'N/A') == 'N/A' and extracted.get('ARRIVAL_SUBJECT', 'N/A') != 'N/A':
        extracted['ARRIVAL'] = extracted['ARRIVAL_SUBJECT']
    
    # Calculate TDF as nights × 20
    try:
        nights = extracted.get('NIGHTS', 'N/A')
        if nights != 'N/A' and str(nights).isdigit():
            nights_num = int(nights)
            tdf_amount = nights_num * 20
            extracted['TDF'] = str(tdf_amount)
            extracted['TDF_AED'] = f"AED {tdf_amount:,.2f}"
        else:
            extracted['TDF'] = "N/A"
    except:
        extracted['TDF'] = "N/A"
    
    # Calculate ADR (Average Daily Rate) = NET_TOTAL / NIGHTS
    try:
        net_total = extracted.get('NET_TOTAL', 'N/A')
        nights = extracted.get('NIGHTS', 'N/A')
        if (net_total != 'N/A' and nights != 'N/A' and 
            str(nights).isdigit() and str(net_total).replace(',', '').replace('.', '').isdigit()):
            nights_num = int(nights)
            net_total_num = float(str(net_total).replace(',', ''))
            if nights_num > 0:
                adr = net_total_num / nights_num
                extracted['ADR'] = f"{adr:.2f}"
                extracted['ADR_AED'] = f"AED {adr:,.2f}"
            else:
                extracted['ADR'] = "N/A"
        else:
            extracted['ADR'] = "N/A"
    except:
        extracted['ADR'] = "N/A"
    
    # Set AMOUNT = NET_TOTAL for consistency
    try:
        net_total = extracted.get('NET_TOTAL', 'N/A')
        if net_total != 'N/A':
            amount_num = float(str(net_total).replace(',', ''))
            extracted['AMOUNT'] = net_total
            extracted['AMOUNT_AED'] = f"AED {amount_num:,.2f}"
        else:
            extracted['AMOUNT'] = "N/A"
    except:
        extracted['AMOUNT'] = "N/A"
    
    return extracted

def get_current_mailbox_info(outlook, namespace):
    """Get information about the current active mailbox"""
    try:
        # Get the active explorer (current Outlook window)
        explorer = outlook.ActiveExplorer()
        if explorer:
            current_folder = explorer.CurrentFolder()
            print(f"Current folder: {current_folder.Name}")
            print(f"Current folder path: {current_folder.FolderPath}")
            
            # Try to get the parent store (mailbox)
            store = current_folder.Store
            print(f"Current store/mailbox: {store.DisplayName}")
            
            return current_folder, store
        else:
            print("[INFO] No active Outlook window found, using default inbox")
            return None, None
    except Exception as e:
        print(f"[INFO] Could not get current mailbox info: {e}")
        return None, None

def search_all_folders_in_mailbox(store, days=2):
    """Search all folders in the current mailbox"""
    print(f"[INFO] Searching all folders in mailbox: {store.DisplayName}")
    
    all_matching_emails = []
    folders_searched = 0
    total_emails = 0
    
    def search_folder_recursive(folder, depth=0):
        nonlocal folders_searched, total_emails, all_matching_emails
        
        indent = "  " * depth
        try:
            # Skip system folders that might cause issues
            folder_name = folder.Name.lower()
            if any(skip in folder_name for skip in ['calendar', 'contacts', 'tasks', 'notes', 'journal']):
                print(f"{indent}[SKIP] Skipping {folder.Name} (system folder)")
                return
            
            folders_searched += 1
            print(f"{indent}[SEARCH] Folder: {folder.Name} ({folder.FolderPath})")
            
            # Get items in this folder
            items = folder.Items
            folder_count = len(items)
            total_emails += folder_count
            
            print(f"{indent}  - Found {folder_count} items")
            
            if folder_count > 0:
                # Apply date filter
                since_date = (datetime.now() - timedelta(days=days)).strftime("%m/%d/%Y")
                try:
                    filtered_items = items.Restrict(f'[ReceivedTime] >= "{since_date}" OR [SentOn] >= "{since_date}"')
                    filtered_count = len(filtered_items)
                    print(f"{indent}  - {filtered_count} items in last {days} days")
                    
                    # Search through filtered items
                    matches_in_folder = search_items_in_folder(filtered_items, folder.Name)
                    all_matching_emails.extend(matches_in_folder)
                    
                    if matches_in_folder:
                        print(f"{indent}  - [MATCH] Found {len(matches_in_folder)} matching emails")
                        
                except Exception as e:
                    print(f"{indent}  - [ERROR] Could not filter items: {e}")
            
            # Search subfolders
            if folder.Folders.Count > 0:
                print(f"{indent}  - Searching {folder.Folders.Count} subfolders...")
                for subfolder in folder.Folders:
                    search_folder_recursive(subfolder, depth + 1)
                    
        except Exception as e:
            print(f"{indent}[ERROR] Error searching folder {folder.Name}: {e}")
    
    # Start recursive search from the root folder of the store
    try:
        root_folder = store.GetRootFolder()
        search_folder_recursive(root_folder)
    except Exception as e:
        print(f"[ERROR] Could not access root folder: {e}")
    
    print(f"\n[SUMMARY] Search completed:")
    print(f"  - Folders searched: {folders_searched}")
    print(f"  - Total emails found: {total_emails}")
    print(f"  - Matching emails: {len(all_matching_emails)}")
    
    return all_matching_emails

def search_items_in_folder(items, folder_name):
    """Search for matching items in a specific folder"""
    matching_emails = []
    
    for item in items:
        try:
            # Check if this is an email item
            if not hasattr(item, 'Subject'):
                continue
            
            # Get email properties
            sender_email = getattr(item, 'SenderEmailAddress', '') or ''
            sender_name = getattr(item, 'SenderName', '') or ''
            subject = getattr(item, 'Subject', '') or ''
            body = getattr(item, 'Body', '') or ''
            received_time = getattr(item, 'ReceivedTime', '') or getattr(item, 'SentOn', '')
            
            # Check if this email matches our criteria
            email_text = (subject + ' ' + body + ' ' + sender_email + ' ' + sender_name).lower()
            
            # Determine match type
            match_type = None
            if 'reservations.gmhd@millenniumhotels.com' in sender_email.lower():
                match_type = 'Millennium Hotels Reservations'
            elif 'avital' in email_text:
                match_type = 'Avital Boaz'
            elif 'shi guang' in email_text or 'shi' in email_text:
                match_type = 'Shi Guang'
            
            if match_type:
                email_info = {
                    'folder': folder_name,
                    'match_type': match_type,
                    'subject': subject,
                    'sender_email': sender_email,
                    'sender_name': sender_name,
                    'received_time': received_time,
                    'has_attachments': hasattr(item, 'Attachments') and item.Attachments.Count > 0,
                    'attachment_count': item.Attachments.Count if hasattr(item, 'Attachments') else 0,
                    'body_snippet': body[:300] + '...' if len(body) > 300 else body,
                    'extracted_data': {},
                    'pdf_attachments': []
                }
                
                # For noreply-reservations emails, extract data from the email body and subject
                if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
                    # Combine subject and body for extraction
                    full_content = subject + "\n" + body
                    extracted_fields = extract_reservation_fields(full_content, sender_email)
                    email_info['extracted_data'] = extracted_fields
                    
                    # Format currency fields
                    for field in ['NET_TOTAL', 'TDF']:
                        if extracted_fields.get(field) != 'N/A' and extracted_fields.get(field):
                            try:
                                amount = float(str(extracted_fields[field]).replace(',', ''))
                                extracted_fields[f'{field}_AED'] = f"AED {amount:,.2f}"
                            except:
                                pass
                
                # Process PDF attachments if present
                if email_info['has_attachments']:
                    for attachment in item.Attachments:
                        filename = getattr(attachment, 'FileName', '')
                        
                        if filename and filename.lower().endswith('.pdf'):
                            try:
                                # Save attachment temporarily
                                temp_path = f"temp_{filename.replace(' ', '_').replace('/', '_')}"
                                attachment.SaveAsFile(temp_path)
                                
                                with open(temp_path, 'rb') as f:
                                    pdf_data = f.read()
                                    text = extract_pdf_text(pdf_data)
                                    
                                    if text:
                                        extracted_fields = extract_reservation_fields(text, sender_email)
                                        
                                        # Format currency fields
                                        for field in ['NET_TOTAL', 'TDF']:
                                            if extracted_fields.get(field) != 'N/A':
                                                try:
                                                    amount = float(extracted_fields[field].replace(',', ''))
                                                    extracted_fields[f'{field}_AED'] = f"AED {amount:,.2f}"
                                                except:
                                                    pass
                                        
                                        email_info['extracted_data'] = extracted_fields
                                        email_info['pdf_attachments'].append({
                                            'filename': filename,
                                            'text_sample': text[:200] + '...' if len(text) > 200 else text
                                        })
                                
                                # Clean up temp file
                                if os.path.exists(temp_path):
                                    os.remove(temp_path)
                                    
                            except Exception as e:
                                print(f"         [FAIL] Error processing PDF {filename}: {e}")
                        else:
                            email_info['pdf_attachments'].append({
                                'filename': filename,
                                'type': 'non-pdf'
                            })
                
                matching_emails.append(email_info)
                
        except Exception as e:
            continue  # Skip problematic items
    
    return matching_emails

def search_current_mailbox():
    """Search for emails in the current mailbox"""
    print("=== Current Mailbox Email Search (Last 2 Days) ===")
    print()
    
    # Connect to Outlook
    print("1. Connecting to Outlook...")
    outlook, namespace = connect_to_outlook()
    if not outlook or not namespace:
        print("   [FAIL] Failed to connect to Outlook")
        return
    print("   [OK] Connected to Outlook successfully")
    print()
    
    # Get current mailbox info
    print("2. Getting current mailbox information...")
    current_folder, store = get_current_mailbox_info(outlook, namespace)
    
    if not store:
        print("   [INFO] Using default mailbox")
        # Fallback to default store
        store = namespace.GetDefaultFolder(6).Store  # Default inbox store
    
    print()
    
    # Search all folders in the current mailbox
    print("3. Searching current mailbox for target emails...")
    print("   Target criteria:")
    print("   - Sender: reservations.gmhd@millenniumhotels.com")
    print("   - Content containing 'Avital' or 'Shi Guang'")
    print("   - Last 2 days")
    print()
    
    all_matching_emails = search_all_folders_in_mailbox(store, days=2)
    
    # Display results
    print()
    print("4. [RESULTS] Search Results:")
    
    if all_matching_emails:
        # Group by match type
        match_counts = {}
        for email in all_matching_emails:
            match_type = email['match_type']
            match_counts[match_type] = match_counts.get(match_type, 0) + 1
        
        print(f"   Total matching emails found: {len(all_matching_emails)}")
        for match_type, count in match_counts.items():
            print(f"   - {match_type}: {count} emails")
        
        print()
        print("5. Detailed Email Results:")
        print()
        
        # Specified fields to display (as per Entered On sheet requirements)
        specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                          'NIGHTS', 'PERSONS', 'ROOM', 'TDF', 'NET_TOTAL', 
                          'ADR', 'AMOUNT', 'RATE_CODE', 'COMPANY']
        
        for i, email in enumerate(all_matching_emails):
            print(f"=== Email {i+1} ({email['match_type']}) ===")
            print(f"Folder: {email['folder']}")
            print(f"Subject: {email['subject']}")
            print(f"From: {email['sender_name']} ({email['sender_email']})")
            print(f"Received: {email['received_time']}")
            print(f"Attachments: {email['attachment_count']}")
            
            if email['body_snippet'].strip():
                print(f"Body Preview:")
                print(f"   {email['body_snippet']}")
            
            # Show extracted data if available
            if email['extracted_data']:
                print(f"Extracted Reservation Fields:")
                for field in specified_fields:
                    value = email['extracted_data'].get(field, 'N/A')
                    print(f"   {field}: {value}")
                
                # Show formatted currency if available
                print(f"\nFormatted Currency Fields:")
                for field in ['TDF_AED', 'NET_TOTAL_AED', 'ADR_AED', 'AMOUNT_AED']:
                    if field in email['extracted_data'] and email['extracted_data'][field] != 'N/A':
                        original_field = field.replace('_AED', '')
                        print(f"   {original_field}_FORMATTED: {email['extracted_data'][field]}")
                
                # Show PDF content sample
                if email['pdf_attachments']:
                    for pdf_info in email['pdf_attachments']:
                        if 'text_sample' in pdf_info:
                            print(f"PDF Content Sample ({pdf_info['filename']}):")
                            print(f"   {pdf_info['text_sample']}")
            else:
                print(f"No reservation data extracted")
                if email['pdf_attachments']:
                    print(f"Attachments found:")
                    for pdf_info in email['pdf_attachments']:
                        print(f"   - {pdf_info['filename']}")
            
            print()
            print("-" * 80)
            print()
    else:
        print("   [NONE] No matching emails found")
        print("   This could mean:")
        print("   - No emails from the target senders in the last 2 days")
        print("   - The search criteria need adjustment")

if __name__ == "__main__":
    search_current_mailbox()