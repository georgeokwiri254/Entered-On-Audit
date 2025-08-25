"""
Extract mail fields from Dubai Link MSG file
"""

import os
import sys
import pandas as pd
import re
from datetime import datetime
import win32com.client
import pythoncom

# Add the current directory to sys.path to import our modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the regex patterns and extraction functions from streamlit_app
from streamlit_app import extract_reservation_fields, NOREPLY_PATTERNS, DEFAULT_PATTERNS

def read_msg_file(msg_path):
    """Read .msg file using Outlook COM"""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Open the .msg file
        msg = outlook.Session.OpenSharedItem(msg_path)
        
        # Extract email properties
        email_data = {
            'subject': getattr(msg, 'Subject', ''),
            'sender': getattr(msg, 'SenderEmailAddress', ''),
            'sender_name': getattr(msg, 'SenderName', ''),
            'body': getattr(msg, 'Body', ''),
            'received_time': str(getattr(msg, 'ReceivedTime', '')),
            'attachments': []
        }
        
        # Process attachments if any
        if hasattr(msg, 'Attachments') and msg.Attachments.Count > 0:
            for attachment in msg.Attachments:
                filename = getattr(attachment, 'FileName', '')
                email_data['attachments'].append({
                    'filename': filename,
                    'type': 'pdf' if filename.lower().endswith('.pdf') else 'other'
                })
        
        return email_data
        
    except Exception as e:
        print(f"Error reading .msg file: {e}")
        return None
    finally:
        pythoncom.CoUninitialize()

def extract_dubai_link_fields(msg_path):
    """Extract required fields from Dubai Link MSG file"""
    
    print(f"Extracting data from: {msg_path}")
    print("="*80)
    
    # Read the .msg file
    email_data = read_msg_file(msg_path)
    
    if not email_data:
        print("Failed to read .msg file")
        return None
    
    print(f"Email Subject: {email_data['subject']}")
    print(f"Sender: {email_data['sender']} ({email_data['sender_name']})")
    
    # Manual extraction for Dubai Link format
    body = email_data['body']
    
    # Extract specific fields using regex patterns for Dubai Link
    import re
    
    # Extract names - Dubai Link specific mapping
    first_name_match = re.search(r'Name:\s*([A-Z]+)', body)
    last_name_match = re.search(r'Last Name:\s*([A-Z]+)', body)
    
    # For Dubai Link: MAIL_FIRST_NAME = Name field, MAIL_FULL_NAME = Last Name field
    first_name = first_name_match.group(1) if first_name_match else 'N/A'  # SOHEIL
    full_name = last_name_match.group(1) if last_name_match else 'N/A'     # RADIOM
    
    # Extract dates
    arrival_match = re.search(r'Arrival Date:\s*(\d{2}/\d{2}/\d{4})', body)
    departure_match = re.search(r'Departure Date:\s*(\d{2}/\d{2}/\d{4})', body)
    
    arrival = arrival_match.group(1) if arrival_match else 'N/A'
    departure = departure_match.group(1) if departure_match else 'N/A'
    
    # Calculate nights
    nights = 1  # Default from the example
    if arrival != 'N/A' and departure != 'N/A':
        try:
            from datetime import datetime
            arr_date = datetime.strptime(arrival, '%d/%m/%Y')
            dep_date = datetime.strptime(departure, '%d/%m/%Y')
            nights = (dep_date - arr_date).days
        except:
            nights = 1
    
    # Extract persons
    persons_match = re.search(r'\((\d+) Adult\)', body)
    persons = int(persons_match.group(1)) if persons_match else 1
    
    # Extract room type
    room_match = re.search(r'(\d+ x [^(]+\([^)]+\)[^)]+)', body)
    room = room_match.group(1).strip() if room_match else 'N/A'
    
    # Extract promo code (rate code)
    promo_match = re.search(r'Promo code:\s*([A-Z0-9{}\s]+)', body)
    rate_code = promo_match.group(1).strip() if promo_match else 'N/A'
    
    # Extract booking cost (net total)
    cost_match = re.search(r'Booking cost price:\s*([\d,.]+)\s*AED', body)
    net_total = float(cost_match.group(1).replace(',', '')) if cost_match else 0
    
    # Calculate TDF based on room type and nights
    tdf = 0
    if room != 'N/A':
        is_two_bedroom = '2BA' in room.upper() or 'Two Bedroom' in room
        tdf_rate = 40 if is_two_bedroom else 20
        
        # For 30+ nights, use 30 as the multiplier instead of actual nights
        effective_nights = min(nights, 30) if nights >= 30 else nights
        tdf = effective_nights * tdf_rate
    
    # Calculate derived values
    mail_total = net_total + tdf if net_total > 0 else 0
    mail_amount = net_total / 1.225 if net_total > 0 else 0
    mail_adr = mail_amount / nights if nights > 0 and mail_amount > 0 else 0
    
    # Map to mail variables
    mail_vars = {
        'MAIL_FIRST_NAME': first_name,
        'MAIL_FULL_NAME': full_name,
        'MAIL_ARRIVAL': arrival,
        'MAIL_DEPARTURE': departure,
        'MAIL_NIGHTS': nights,
        'MAIL_PERSONS': persons,
        'MAIL_ROOM': room,
        'MAIL_RATE_CODE': rate_code,
        'MAIL_C_T_S': 'Dubai Link',  # Travel agency name
        'MAIL_NET_TOTAL': net_total if net_total > 0 else 'N/A',
        'MAIL_TDF': tdf if tdf > 0 else 'N/A',
        'MAIL_TOTAL': mail_total if mail_total > 0 else 'N/A',
        'MAIL_AMOUNT': mail_amount if mail_amount > 0 else 'N/A',
        'MAIL_ADR': mail_adr if mail_adr > 0 else 'N/A'
    }
    
    # Display results
    print(f"\nExtracted Mail Variables:")
    print("="*80)
    
    for var_name, value in mail_vars.items():
        if isinstance(value, float):
            formatted_value = f"AED {value:,.2f}"
        else:
            formatted_value = value
        print(f"{var_name:20}: {formatted_value}")
    
    # Also print the email body for manual verification
    print(f"\n\nEmail Body (for manual verification):")
    print("-" * 80)
    try:
        body_text = email_data['body'][:2000] + "..." if len(email_data['body']) > 2000 else email_data['body']
        # Remove or replace problematic Unicode characters
        body_text = body_text.encode('ascii', 'ignore').decode('ascii')
        print(body_text)
    except Exception as e:
        print(f"Could not display email body due to encoding: {e}")
        print("Body contains non-ASCII characters")
    print("-" * 80)
    
    return mail_vars

if __name__ == "__main__":
    # Extract from the Dubai Link MSG file
    msg_file_path = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\Travel Agency TO\Dubai Link\Confirmed Booking with Ref. No. VF1F41.msg"
    
    if os.path.exists(msg_file_path):
        mail_variables = extract_dubai_link_fields(msg_file_path)
        print(f"\nExtraction completed!")
    else:
        print(f"File not found: {msg_file_path}")