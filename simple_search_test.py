"""
Simple test script to search for Avital Boaz
"""
import pandas as pd
import os
from entered_on_converter import process_entered_on_report

def test_avital_search():
    print("=== Testing Avital Boaz Search ===")
    print()
    
    # 1. Get the latest file manually
    base_path = "P:\\Reservation\\Entered on"
    if not os.path.exists(base_path):
        print(f"Error: Base path does not exist: {base_path}")
        return
    
    # Get all directories
    directories = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    if not directories:
        print("No directories found")
        return
    
    # Sort by modification time (latest first)
    directories.sort(key=lambda x: os.path.getmtime(os.path.join(base_path, x)), reverse=True)
    latest_dir = directories[0]
    latest_dir_path = os.path.join(base_path, latest_dir)
    
    # Get xlsm files, skip temp files
    xlsm_files = [f for f in os.listdir(latest_dir_path) 
                 if f.lower().endswith('.xlsm') and not f.startswith('~$')]
    
    if not xlsm_files:
        print(f"No .xlsm files found in {latest_dir}")
        return
    
    xlsm_files.sort(key=lambda x: os.path.getmtime(os.path.join(latest_dir_path, x)), reverse=True)
    latest_file = xlsm_files[0]
    latest_file_path = os.path.join(latest_dir_path, latest_file)
    
    print(f"1. Using file: {latest_dir}\\{latest_file}")
    print()
    
    # 2. Process the Excel file
    print("2. Processing Excel file...")
    try:
        processed_df, csv_path = process_entered_on_report(latest_file_path)
        print(f"   Processed {len(processed_df)} records")
        print(f"   Columns: {list(processed_df.columns)}")
        print()
    except Exception as e:
        print(f"Error processing file: {e}")
        return
    
    # 3. Search for Avital
    print("3. Searching for 'Avital' in the data...")
    
    # Search in FULL_NAME
    avital_records = processed_df[processed_df['FULL_NAME'].str.contains('Avital', case=False, na=False)]
    
    if avital_records.empty:
        print("   No records found in FULL_NAME column")
        
        # Try other columns
        search_columns = ['FIRST_NAME', 'COMPANY', 'COMPANY_CLEAN'] if 'FIRST_NAME' in processed_df.columns else ['COMPANY', 'COMPANY_CLEAN']
        
        for col in search_columns:
            if col in processed_df.columns:
                matches = processed_df[processed_df[col].str.contains('Avital', case=False, na=False)]
                if not matches.empty:
                    print(f"   Found {len(matches)} records with 'Avital' in {col} column")
                    avital_records = matches
                    break
        
        if avital_records.empty:
            print("   No records found with 'Avital' in any column")
            print("   Sample guest names:")
            for i, name in enumerate(processed_df['FULL_NAME'].dropna().head(5)):
                print(f"     - {name}")
            return
    else:
        print(f"   Found {len(avital_records)} records with 'Avital' in FULL_NAME")
    
    # 4. Show the record details
    print()
    print("4. Record details for Avital:")
    
    specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                      'NIGHTS', 'PERSONS', 'ROOM', 'TDF', 'NET_TOTAL', 
                      'RATE_CODE', 'COMPANY']
    
    for idx, (row_idx, record) in enumerate(avital_records.iterrows()):
        print(f"")
        print(f"--- Record {idx + 1} (Row {row_idx}) ---")
        
        for field in specified_fields:
            if field in record:
                value = record[field]
                if pd.isna(value):
                    value = 'N/A'
                elif field in ['TDF', 'NET_TOTAL'] and pd.notna(value):
                    try:
                        amount = float(str(value).replace(',', ''))
                        value = f"AED {amount:,.2f}"
                    except:
                        pass
                print(f"  {field}: {value}")
            else:
                print(f"  {field}: Column not found")
        
        # Show additional available fields
        print("  ")
        print("  Additional fields available:")
        other_fields = [col for col in record.index if col not in specified_fields and pd.notna(record[col])]
        for field in other_fields[:5]:  # Show first 5 additional fields
            value = record[field]
            if isinstance(value, (int, float)) and field in ['AMOUNT', 'ADR']:
                value = f"AED {value:,.2f}"
            print(f"    {field}: {value}")
    
    print()
    print("=== Search Complete ===")

if __name__ == "__main__":
    test_avital_search()