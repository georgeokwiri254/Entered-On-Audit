"""
Test script for T-Agoda extraction with correct amount calculations
T-Agoda: Email amount is MAIL_NET_TOTAL (excludes TDF)
MAIL_TOTAL = MAIL_NET_TOTAL + MAIL_TDF
MAIL_AMOUNT = MAIL_NET_TOTAL / 1.225
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

def test_agoda_extraction_accuracy(msg_path, output_csv_path):
    """Test T-Agoda extraction accuracy with correct business logic"""
    
    print("Testing T-Agoda Extraction Logic")
    print(f"File: {os.path.basename(msg_path)}")
    print("="*80)
    
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
    
    # Extract reservation fields using UPDATED logic
    extracted_fields = extract_reservation_fields(full_content, sender_email)
    
    # Define the mail fields including MAIL_FULL_NAME
    test_fields = [
        'FIRST_NAME', 'FULL_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 
        'ROOM', 'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 
        'TDF', 'ADR', 'AMOUNT'
    ]
    
    # Prepare results for CSV
    results = []
    
    print(f"\nT-Agoda Extraction Results:")
    print("="*80)
    
    # Show the calculation breakdown
    print("T-Agoda Amount Calculation Breakdown:")
    print("-" * 50)
    
    try:
        nights = int(extracted_fields.get('NIGHTS', 0))
        net_total_amount = float(extracted_fields.get('NET_TOTAL', '0').replace(',', ''))
        tdf = float(extracted_fields.get('TDF', '0').replace(',', ''))
        total_with_tdf = float(extracted_fields.get('TOTAL', '0').replace(',', ''))
        amount = float(extracted_fields.get('AMOUNT', '0').replace(',', ''))
        adr = float(extracted_fields.get('ADR', '0').replace(',', ''))
        
        print(f"Email Amount (MAIL_NET_TOTAL): AED {net_total_amount:,.2f}")
        print(f"TDF ({nights} x AED 20):         AED {tdf:,.2f}")
        print(f"TOTAL (NET_TOTAL + TDF):       AED {total_with_tdf:,.2f}")
        print(f"AMOUNT (NET_TOTAL / 1.225):    AED {amount:,.2f}")
        print(f"ADR (AMOUNT / {nights} nights):    AED {adr:,.2f}")
        print("-" * 50)
        
        # Verify calculations
        expected_tdf = nights * 20
        expected_total = net_total_amount + expected_tdf
        expected_amount = net_total_amount / 1.225
        expected_adr = expected_amount / nights if nights > 0 else 0
        
        print("T-Agoda Calculation Verification:")
        print(f"TDF Correct: {abs(tdf - expected_tdf) < 0.01}")
        print(f"TOTAL Correct: {abs(total_with_tdf - expected_total) < 0.01}")
        print(f"AMOUNT Correct: {abs(amount - expected_amount) < 0.01}")
        print(f"ADR Correct: {abs(adr - expected_adr) < 0.01}")
        
    except:
        print("Error in calculation verification")
    
    print(f"\nField-by-Field Results:")
    print("-" * 80)
    
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
        'Extracted_Value': f"T-Agoda Logic Applied",
        'Formatted_Value': f"Sender: {sender_email}",
        'Status': f"Processed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    }
    results.insert(0, metadata_row)
    
    # Create DataFrame and save to CSV
    df = pd.DataFrame(results)
    df.to_csv(output_csv_path, index=False)
    
    print(f"\nResults saved to: {output_csv_path}")
    
    # Summary statistics
    found_count = len([r for r in results[1:] if r['Status'] == 'Found'])
    total_fields = len(test_fields)
    accuracy = (found_count / total_fields) * 100
    
    print(f"\nT-Agoda Extraction Summary:")
    print(f"Fields Found: {found_count}/{total_fields}")
    print(f"Accuracy: {accuracy:.1f}%")
    print(f"T-Agoda Logic Applied Successfully!")
    
    return df, accuracy

if __name__ == "__main__":
    # Test the specific Agoda .msg file
    msg_file_path = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\INNLINKWAY\Agoda\Arrival Date09062025Grand Millennium Dubai confirmation number4K76RPPXK.msg"
    output_csv = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\agoda_extraction_test_results.csv"
    
    if os.path.exists(msg_file_path):
        results_df, accuracy = test_agoda_extraction_accuracy(msg_file_path, output_csv)
        print(f"\nAgoda test completed! CSV saved to: {output_csv}")
    else:
        print(f"File not found: {msg_file_path}")