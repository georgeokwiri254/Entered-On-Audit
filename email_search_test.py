"""
Test script to search Outlook emails for Avital Boaz reservation data
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

def search_emails_for_avital():
    print("=== Searching Outlook Emails for Avital Boaz ===")
    print()
    
    # Connect to Outlook
    print("1. Connecting to Outlook...")
    outlook, namespace = connect_to_outlook()
    if not outlook or not namespace:
        print("   Failed to connect to Outlook")
        return
    print("   Connected to Outlook successfully")
    print()
    
    # Get inbox
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        print(f"2. Accessing inbox: {inbox.Name}")
    except Exception as e:
        print(f"   Error accessing inbox: {e}")
        return
    
    # Search for emails in the last 30 days
    since_date = (datetime.now() - timedelta(days=30)).strftime("%m/%d/%Y")
    print(f"3. Searching emails since: {since_date}")
    
    try:
        # Search for emails containing "Avital"
        messages = inbox.Items.Restrict(f'[ReceivedTime] >= "{since_date}"')
        print(f"   Found {len(messages)} total emails in date range")
        
        matching_emails = []
        processed_count = 0
        
        print("4. Searching for 'Avital' in emails...")
        
        for message in messages:
            processed_count += 1
            if processed_count % 50 == 0:
                print(f"   Processed {processed_count} emails...")
            
            try:
                if not hasattr(message, 'Subject'):
                    continue
                
                # Check sender - look for reservations.gmhd@millenniumhotels.com
                sender = getattr(message, 'SenderEmailAddress', '') or ''
                sender_name = getattr(message, 'SenderName', '') or ''
                
                # Get email content
                subject = getattr(message, 'Subject', '') or ''
                body = getattr(message, 'Body', '') or ''
                received_time = getattr(message, 'ReceivedTime', '')
                
                # Check if this email contains "Avital"
                email_text = (subject + ' ' + body + ' ' + sender + ' ' + sender_name).lower()
                
                if 'avital' in email_text:
                    print(f"   >> Found email with 'Avital'!")
                    print(f"     Subject: {subject}")
                    print(f"     From: {sender_name} ({sender})")
                    print(f"     Received: {received_time}")
                    
                    email_info = {
                        'subject': subject,
                        'sender': sender,
                        'sender_name': sender_name,
                        'received_time': received_time,
                        'attachments': [],
                        'extracted_data': {},
                        'body_snippet': body[:200] + '...' if len(body) > 200 else body
                    }
                    
                    # Process attachments (PDFs)
                    if hasattr(message, 'Attachments') and message.Attachments.Count > 0:
                        print(f"     Processing {message.Attachments.Count} attachments...")
                        
                        for attachment in message.Attachments:
                            filename = getattr(attachment, 'FileName', '')
                            print(f"       - {filename}")
                            
                            if filename and filename.lower().endswith('.pdf'):
                                try:
                                    # Save attachment temporarily
                                    temp_path = f"temp_{filename}"
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
                                            email_info['pdf_text_snippet'] = text[:500] + '...' if len(text) > 500 else text
                                            
                                            print(f"       >> Extracted data from PDF")
                                        else:
                                            print(f"       >> Could not extract text from PDF")
                                    
                                    # Clean up temp file
                                    if os.path.exists(temp_path):
                                        os.remove(temp_path)
                                        
                                except Exception as e:
                                    print(f"       >> Error processing PDF {filename}: {e}")
                    
                    matching_emails.append(email_info)
                    
            except Exception as e:
                print(f"   Error processing message: {e}")
                continue
        
        print(f"")
        print(f"5. Search Results Summary:")
        print(f"   Total emails processed: {processed_count}")
        print(f"   Emails matching 'Avital': {len(matching_emails)}")
        print()
        
        # Display results
        if matching_emails:
            print("6. Detailed Results for Avital Boaz:")
            print()
            
            specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                              'NIGHTS', 'PERSONS', 'ROOM', 'TDF', 'NET_TOTAL', 
                              'RATE_CODE', 'COMPANY']
            
            for i, email in enumerate(matching_emails):
                print(f"--- Email {i+1} ---")
                print(f"Subject: {email['subject']}")
                print(f"From: {email['sender_name']} ({email['sender']})")
                print(f"Received: {email['received_time']}")
                print(f"Body snippet: {email['body_snippet']}")
                print()
                
                if email['extracted_data']:
                    print("Extracted Reservation Data:")
                    for field in specified_fields:
                        value = email['extracted_data'].get(field, 'N/A')
                        print(f"  {field}: {value}")
                    
                    # Show formatted currency if available
                    if 'TDF_AED' in email['extracted_data']:
                        print(f"  TDF_FORMATTED: {email['extracted_data']['TDF_AED']}")
                    if 'NET_TOTAL_AED' in email['extracted_data']:
                        print(f"  NET_TOTAL_FORMATTED: {email['extracted_data']['NET_TOTAL_AED']}")
                    
                    print()
                    if 'pdf_text_snippet' in email:
                        print("PDF Content Sample:")
                        print(f"  {email['pdf_text_snippet']}")
                else:
                    print(">> No reservation data extracted from this email")
                
                print()
                print("-" * 50)
                print()
        else:
            print(">> No emails found containing 'Avital'")
            print("   This could mean:")
            print("   - No emails from reservations.gmhd@millenniumhotels.com contain 'Avital'")
            print("   - The emails are older than 30 days")
            print("   - The search term needs to be adjusted")
        
    except Exception as e:
        print(f"Error during email search: {e}")

if __name__ == "__main__":
    search_emails_for_avital()