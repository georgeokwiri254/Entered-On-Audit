"""
Test script to search for Avital Boaz and show email extraction results
"""
import pandas as pd
import sys
import os
from entered_on_converter import process_entered_on_report
from streamlit_app import (
    connect_to_outlook, 
    search_emails_for_reservation,
    get_latest_file_from_path
)

def test_avital_search():
    print("=== Testing Avital Boaz Email Search ===\n")
    
    # 1. Get the latest file
    print("1. Getting latest Excel file...")
    latest_file, status_msg = get_latest_file_from_path()
    if not latest_file:
        print(f"Error: {status_msg}")
        return
    print(f"✅ Found: {status_msg}\n")
    
    # 2. Process the Excel file
    print("2. Processing Excel file...")
    try:
        processed_df, csv_path = process_entered_on_report(latest_file)
        print(f"✅ Processed {len(processed_df)} records\n")
    except Exception as e:
        print(f"Error processing file: {e}")
        return
    
    # 3. Find Avital Boaz in the data
    print("3. Searching for 'Avital Boaz' in the processed data...")
    avital_records = processed_df[processed_df['FULL_NAME'].str.contains('Avital', case=False, na=False)]
    
    if avital_records.empty:
        print("❌ No records found for 'Avital' in FULL_NAME column")
        
        # Try other columns
        print("Searching in other columns...")
        for col in ['FIRST_NAME', 'COMPANY', 'COMPANY_CLEAN']:
            if col in processed_df.columns:
                matches = processed_df[processed_df[col].str.contains('Avital', case=False, na=False)]
                if not matches.empty:
                    print(f"✅ Found {len(matches)} records with 'Avital' in {col} column")
                    avital_records = matches
                    break
        
        if avital_records.empty:
            print("❌ No records found with 'Avital' in any column")
            print("Available guest names (first 10):")
            for i, name in enumerate(processed_df['FULL_NAME'].dropna().head(10)):
                print(f"  - {name}")
            return
    else:
        print(f"✅ Found {len(avital_records)} records with 'Avital' in FULL_NAME")
    
    # 4. Show the record details
    print("\n4. Record details for Avital:")
    for idx, record in avital_records.iterrows():
        print(f"\n--- Record {idx + 1} ---")
        specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                          'NIGHTS', 'PERSONS', 'ROOM', 'TDF', 'NET_TOTAL', 
                          'RATE_CODE', 'COMPANY']
        
        for field in specified_fields:
            value = record.get(field, 'N/A')
            if pd.isna(value):
                value = 'N/A'
            print(f"  {field}: {value}")
    
    # 5. Test email search
    print("\n5. Testing email search for Avital...")
    try:
        outlook, namespace = connect_to_outlook()
        if outlook and namespace:
            print("✅ Connected to Outlook")
            
            # Use the first Avital record for email search
            first_record = avital_records.iloc[0].to_dict()
            print(f"Searching emails for: {first_record.get('FULL_NAME', 'Unknown')}")
            
            matching_emails = search_emails_for_reservation(outlook, namespace, first_record, days=14)
            print(f"✅ Found {len(matching_emails)} matching emails")
            
            if matching_emails:
                for i, email in enumerate(matching_emails):
                    print(f"\n--- Email {i+1} ---")
                    print(f"Subject: {email['subject']}")
                    print(f"From: {email['sender']}")
                    print(f"Received: {email['received_time']}")
                    print(f"Attachments: {len(email['attachments'])}")
                    
                    if email['extracted_data']:
                        print("\n📄 Extracted Data:")
                        for field in specified_fields:
                            value = email['extracted_data'].get(field, 'N/A')
                            print(f"  {field}: {value}")
                        
                        # Show any additional fields
                        additional_fields = [k for k in email['extracted_data'].keys() 
                                           if k not in specified_fields and email['extracted_data'][k] != 'N/A']
                        if additional_fields:
                            print("\n📄 Additional Extracted Fields:")
                            for field in additional_fields:
                                print(f"  {field}: {email['extracted_data'][field]}")
                    else:
                        print("❌ No data extracted from this email")
            else:
                print("❌ No emails found for Avital")
        else:
            print("❌ Could not connect to Outlook")
    
    except Exception as e:
        print(f"Error during email search: {e}")

if __name__ == "__main__":
    test_avital_search()