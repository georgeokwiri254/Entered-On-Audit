"""
Enhanced Outlook search script to find emails from specific senders in the last two days
Searches for:
1. reservations.gmhd@millenniumhotels.com
2. Avital Boaz emails
Shows results with all field details
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

def extract_reservation_fields(text):
    """Extract reservation fields using regex patterns"""
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
        'TDF': r"TDF[:\s]+(?:AED\s*)?([\\d,]+\.?\\d*)",
    }
    
    extracted = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            # Convert date format to dd/mm/yyyy if it's a date field
            if field in ['ARRIVAL', 'DEPARTURE'] and value != 'N/A':
                try:
                    parsed_date = pd.to_datetime(value, dayfirst=True)
                    extracted[field] = parsed_date.strftime('%d/%m/%Y')
                except:
                    extracted[field] = value
            else:
                extracted[field] = value
        else:
            extracted[field] = "N/A"
    
    return extracted

def search_outlook_emails():
    """Search for emails from specific senders in the last 2 days"""
    print("=== Enhanced Outlook Email Search (Last 2 Days) ===")
    print()
    
    # Connect to Outlook
    print("1. Connecting to Outlook...")
    outlook, namespace = connect_to_outlook()
    if not outlook or not namespace:
        print("   [FAIL] Failed to connect to Outlook")
        return
    print("   [OK] Connected to Outlook successfully")
    print()
    
    # Get inbox
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        print(f"2. Accessing inbox: {inbox.Name}")
    except Exception as e:
        print(f"   [FAIL] Error accessing inbox: {e}")
        return
    
    # Search for emails in the last 2 days
    since_date = (datetime.now() - timedelta(days=2)).strftime("%m/%d/%Y")
    print(f"3. Searching emails since: {since_date}")
    
    # Define search criteria
    target_senders = {
        'reservations.gmhd@millenniumhotels.com': 'Millennium Hotels Reservations',
        'avital': 'Avital Boaz (any email containing Avital)'
    }
    
    try:
        # Get all emails from the last 2 days
        messages = inbox.Items.Restrict(f'[ReceivedTime] >= "{since_date}"')
        print(f"   Found {len(messages)} total emails in date range")
        print()
        
        all_matching_emails = []
        processed_count = 0
        
        print("4. Searching for target senders...")
        print("   Target senders:")
        for sender, description in target_senders.items():
            print(f"   - {sender}: {description}")
        print()
        
        for message in messages:
            processed_count += 1
            if processed_count % 20 == 0:
                print(f"   [INFO] Processed {processed_count} emails...")
            
            try:
                if not hasattr(message, 'Subject'):
                    continue
                
                # Get email properties
                sender_email = getattr(message, 'SenderEmailAddress', '') or ''
                sender_name = getattr(message, 'SenderName', '') or ''
                subject = getattr(message, 'Subject', '') or ''
                body = getattr(message, 'Body', '') or ''
                received_time = getattr(message, 'ReceivedTime', '')
                
                # Check if this email matches our criteria
                email_text = (subject + ' ' + body + ' ' + sender_email + ' ' + sender_name).lower()
                
                # Determine match type
                match_type = None
                if 'reservations.gmhd@millenniumhotels.com' in sender_email.lower():
                    match_type = 'Millennium Hotels Reservations'
                elif 'avital' in email_text:
                    match_type = 'Avital Boaz'
                
                if match_type:
                    print(f"   [MATCH] Found match ({match_type})!")
                    print(f"      Subject: {subject[:60]}{'...' if len(subject) > 60 else ''}")
                    print(f"      From: {sender_name} ({sender_email})")
                    print(f"      Received: {received_time}")
                    
                    email_info = {
                        'match_type': match_type,
                        'subject': subject,
                        'sender_email': sender_email,
                        'sender_name': sender_name,
                        'received_time': received_time,
                        'has_attachments': hasattr(message, 'Attachments') and message.Attachments.Count > 0,
                        'attachment_count': message.Attachments.Count if hasattr(message, 'Attachments') else 0,
                        'body_snippet': body[:300] + '...' if len(body) > 300 else body,
                        'extracted_data': {},
                        'pdf_attachments': []
                    }
                    
                    # Process PDF attachments
                    if email_info['has_attachments']:
                        print(f"      [ATTACH] Processing {email_info['attachment_count']} attachments...")
                        
                        for attachment in message.Attachments:
                            filename = getattr(attachment, 'FileName', '')
                            
                            if filename and filename.lower().endswith('.pdf'):
                                print(f"         [PDF] Processing PDF: {filename}")
                                try:
                                    # Save attachment temporarily
                                    temp_path = f"temp_{filename.replace(' ', '_')}"
                                    attachment.SaveAsFile(temp_path)
                                    
                                    with open(temp_path, 'rb') as f:
                                        pdf_data = f.read()
                                        text = extract_pdf_text(pdf_data)
                                        
                                        if text:
                                            extracted_fields = extract_reservation_fields(text)
                                            
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
                                            
                                            print(f"         [OK] Extracted data from PDF")
                                        else:
                                            print(f"         [FAIL] Could not extract text from PDF")
                                    
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
                    
                    all_matching_emails.append(email_info)
                    print()
                    
            except Exception as e:
                continue  # Skip problematic messages
        
        print(f"5. [SUMMARY] Search Results Summary:")
        print(f"   Total emails processed: {processed_count}")
        print(f"   Total matching emails: {len(all_matching_emails)}")
        
        # Group by match type
        match_counts = {}
        for email in all_matching_emails:
            match_type = email['match_type']
            match_counts[match_type] = match_counts.get(match_type, 0) + 1
        
        for match_type, count in match_counts.items():
            print(f"   - {match_type}: {count} emails")
        print()
        
        # Display detailed results
        if all_matching_emails:
            print("6. [RESULTS] Detailed Email Results:")
            print()
            
            # Specified fields to display
            specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                              'NIGHTS', 'PERSONS', 'ROOM', 'TDF', 'NET_TOTAL', 
                              'RATE_CODE', 'COMPANY']
            
            for i, email in enumerate(all_matching_emails):
                print(f"=== Email {i+1} ({email['match_type']}) ===")
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
                    for field in ['TDF_AED', 'NET_TOTAL_AED']:
                        if field in email['extracted_data']:
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
            print("6. [NONE] No matching emails found")
            print("   This could mean:")
            print("   - No emails from the target senders in the last 2 days")
            print("   - The search criteria need adjustment")
            print("   - The emails are in a different folder")
            print()
            print("   [INFO] Suggestion: Try expanding the date range or check other folders")
        
    except Exception as e:
        print(f"[ERROR] Error during email search: {e}")
        import traceback
        traceback.print_exc()

def show_all_email_fields():
    """Display available email fields for debugging"""
    print("=== Available Email Fields (Debug Mode) ===")
    print()
    
    outlook, namespace = connect_to_outlook()
    if not outlook or not namespace:
        return
    
    try:
        inbox = namespace.GetDefaultFolder(6)
        messages = inbox.Items.Restrict('[ReceivedTime] >= "' + (datetime.now() - timedelta(days=1)).strftime("%m/%d/%Y") + '"')
        
        if len(messages) > 0:
            sample_message = messages[0]
            print("Sample email fields:")
            
            # Common properties to check
            properties = [
                'Subject', 'SenderName', 'SenderEmailAddress', 'ReceivedTime',
                'Body', 'HTMLBody', 'To', 'CC', 'BCC', 'Importance',
                'Size', 'UnRead', 'Categories', 'FlagStatus'
            ]
            
            for prop in properties:
                try:
                    value = getattr(sample_message, prop, 'Not available')
                    if isinstance(value, str) and len(value) > 100:
                        value = value[:100] + '...'
                    print(f"  {prop}: {value}")
                except Exception as e:
                    print(f"  {prop}: Error accessing - {e}")
    
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "debug":
        show_all_email_fields()
    else:
        search_outlook_emails()