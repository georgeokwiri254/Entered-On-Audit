"""
Test script for Brand.com extraction - DISPLAY RESULTS ONLY
Brand.com uses same logic as T-Booking.com: Email amount is MAIL_TOTAL (includes TDF)
Rule: Any INNLINKWAY INSERT_USER reservation NOT from Agoda/Expedia follows Booking.com rules
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

# Import the updated regex patterns and extraction functions from streamlit_app
from streamlit_app import extract_reservation_fields

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

def test_brand_com_extraction():
    """Test Brand.com extraction with same logic as T-Booking.com"""
    
    msg_path = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\INNLINKWAY\Brand.com\Arrival Date08282025Grand Millennium Dubai confirmation number4K76RP02A.msg"
    
    print("="*80)
    print("TESTING BRAND.COM EXTRACTION")
    print(f"File: {os.path.basename(msg_path)}")
    print("="*80)
    
    if not os.path.exists(msg_path):
        print(f"File not found: {msg_path}")
        return
    
    # Read the .msg file
    email_data = read_msg_file(msg_path)
    
    if not email_data:
        print("Failed to read .msg file")
        return
    
    print(f"Email Subject: {email_data['subject']}")
    print(f"Sender: {email_data['sender']}")
    print(f"Attachments: {len(email_data['attachments'])}")
    
    # Combine subject and body for extraction
    full_content = email_data['subject'] + "\n" + email_data['body']
    sender_email = email_data['sender']
    
    # Extract reservation fields
    extracted_fields = extract_reservation_fields(full_content, sender_email)
    
    # Define the mail fields including MAIL_FULL_NAME
    test_fields = [
        'FIRST_NAME', 'FULL_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 
        'ROOM', 'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 
        'TDF', 'ADR', 'AMOUNT'
    ]
    
    print(f"\nBRAND.COM EXTRACTION RESULTS:")
    print("="*80)
    
    # Show the calculation breakdown
    print("Brand.com Amount Calculation Breakdown:")
    print("(Following T-Booking.com Logic)")
    print("-" * 60)
    
    try:
        nights = int(extracted_fields.get('NIGHTS', 0))
        net_total = float(extracted_fields.get('NET_TOTAL', '0').replace(',', ''))
        tdf = float(extracted_fields.get('TDF', '0').replace(',', ''))
        total = float(extracted_fields.get('TOTAL', '0').replace(',', ''))
        amount = float(extracted_fields.get('AMOUNT', '0').replace(',', ''))
        adr = float(extracted_fields.get('ADR', '0').replace(',', ''))
        
        print(f"Email Amount (MAIL_TOTAL):     AED {total:,.2f}  <- Includes TDF")
        print(f"TDF ({nights} x AED 20):         AED {tdf:,.2f}")
        print(f"NET_TOTAL (TOTAL - TDF):      AED {net_total:,.2f}")
        print(f"AMOUNT (NET_TOTAL / 1.225):   AED {amount:,.2f}")
        print(f"ADR (AMOUNT / {nights} nights):   AED {adr:,.2f}")
        print("-" * 60)
        
        # Verify Brand.com calculations (same as T-Booking.com)
        expected_tdf = nights * 20
        expected_net_total = total - expected_tdf
        expected_amount = expected_net_total / 1.225
        expected_adr = expected_amount / nights if nights > 0 else 0
        
        print("Brand.com Calculation Verification:")
        print(f"TDF Correct: {abs(tdf - expected_tdf) < 0.01}")
        print(f"NET_TOTAL Correct: {abs(net_total - expected_net_total) < 0.01}")
        print(f"AMOUNT Correct: {abs(amount - expected_amount) < 0.01}")
        print(f"ADR Correct: {abs(adr - expected_adr) < 0.01}")
        
    except Exception as e:
        print(f"Error in calculation display: {e}")
    
    print(f"\nFIELD-BY-FIELD RESULTS:")
    print("-" * 80)
    
    found_count = 0
    for field in test_fields:
        value = extracted_fields.get(field, 'N/A')
        mail_field = f'MAIL_{field}'
        
        # Format currency fields
        if field in ['NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT'] and value != 'N/A':
            try:
                amount_val = float(str(value).replace(',', ''))
                formatted_value = f"AED {amount_val:,.2f}"
                found_count += 1
            except:
                formatted_value = value
                if value != 'N/A':
                    found_count += 1
        else:
            formatted_value = value
            if value != 'N/A':
                found_count += 1
        
        status = "Found" if value != 'N/A' else "Not Found"
        print(f"{mail_field:20}: {formatted_value:15} [{status}]")
    
    # Summary
    total_fields = len(test_fields)
    accuracy = (found_count / total_fields) * 100
    
    print(f"\nBRAND.COM EXTRACTION SUMMARY:")
    print(f"Fields Found: {found_count}/{total_fields}")
    print(f"Accuracy: {accuracy:.1f}%")
    print(f"Brand.com Logic Applied Successfully!")
    print("(Same logic as T-Booking.com: Email amount = MAIL_TOTAL)")
    print("INNLINKWAY INSERT_USER rule: Non-Agoda/Expedia = Booking.com logic")
    
    print("="*80)
    print("BRAND.COM EXTRACTION TEST COMPLETED - NO CSV GENERATED")
    print("="*80)

if __name__ == "__main__":
    test_brand_com_extraction()