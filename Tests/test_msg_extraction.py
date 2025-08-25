"""
Test script to extract mail fields from .msg file for accuracy testing
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

def test_extraction_accuracy(msg_path, output_csv_path):
    """Test extraction accuracy on specific .msg file"""
    
    print(f"Testing extraction on: {msg_path}")
    print("="*80)
    
    # Read the .msg file
    email_data = read_msg_file(msg_path)
    
    if not email_data:
        print("Failed to read .msg file")
        return
    
    print(f"Email Subject: {email_data['subject']}")
    print(f"Sender: {email_data['sender']} ({email_data['sender_name']})")
    print(f"Attachments: {len(email_data['attachments'])}")
    for att in email_data['attachments']:
        print(f"  - {att['filename']} ({att['type']})")
    
    print("\nEmail Body Preview:")
    print("-" * 50)
    print(email_data['body'][:500] + "..." if len(email_data['body']) > 500 else email_data['body'])
    print("-" * 50)
    
    # Combine subject and body for extraction
    full_content = email_data['subject'] + "\n" + email_data['body']
    sender_email = email_data['sender']
    
    # Extract reservation fields
    extracted_fields = extract_reservation_fields(full_content, sender_email)
    
    # Define the mail fields we want to test
    test_fields = [
        'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 
        'ROOM', 'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 
        'TDF', 'ADR', 'AMOUNT'
    ]
    
    # Prepare results for CSV
    results = []
    
    print(f"\nExtraction Results:")
    print("="*80)
    
    for field in test_fields:
        value = extracted_fields.get(field, 'N/A')
        mail_field = f'MAIL_{field}'
        
        # Format currency fields
        if field in ['NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT'] and value != 'N/A':
            try:
                amount = float(str(value).replace(',', ''))
                formatted_value = f"AED {amount:,.2f}"
            except:
                formatted_value = value
        else:
            formatted_value = value
        
        print(f"{mail_field:20}: {formatted_value}")
        
        results.append({
            'Field': mail_field,
            'Extracted_Value': value,
            'Formatted_Value': formatted_value,
            'Status': 'Found' if value != 'N/A' else 'Not Found'
        })
    
    # Add metadata
    metadata_row = {
        'Field': 'METADATA',
        'Extracted_Value': f"Subject: {email_data['subject'][:100]}...",
        'Formatted_Value': f"Sender: {sender_email}",
        'Status': f"Received: {email_data['received_time']}"
    }
    results.insert(0, metadata_row)
    
    # Create DataFrame and save to CSV
    df = pd.DataFrame(results)
    df.to_csv(output_csv_path, index=False)
    
    print(f"\nResults saved to: {output_csv_path}")
    
    # Summary statistics
    found_count = len([r for r in results[1:] if r['Status'] == 'Found'])  # Exclude metadata
    total_fields = len(test_fields)
    accuracy = (found_count / total_fields) * 100
    
    print(f"\nExtraction Accuracy Summary:")
    print(f"Fields Found: {found_count}/{total_fields}")
    print(f"Accuracy: {accuracy:.1f}%")
    
    return df, accuracy

if __name__ == "__main__":
    # Test the specific .msg file
    msg_file_path = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\INNLINKWAY\Booking.com\Arrival Date09042025Grand Millennium Dubai confirmation number4K76RP0X8.msg"
    output_csv = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\extraction_test_results.csv"
    
    if os.path.exists(msg_file_path):
        results_df, accuracy = test_extraction_accuracy(msg_file_path, output_csv)
        print(f"\nTest completed! CSV saved to: {output_csv}")
    else:
        print(f"File not found: {msg_file_path}")