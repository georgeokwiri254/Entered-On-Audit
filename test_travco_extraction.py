"""
Test script for Travco MSG file extraction
"""

import sys
import os
from pathlib import Path

# Add the Rules directory to path to import parsers
sys.path.append(str(Path(__file__).parent / "Rules" / "Travel Agency TO" / "Travco"))

from travco_parser import extract_travco_fields, is_travco_email

def read_msg_file(file_path):
    """Read MSG file content"""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        return content
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

def main():
    # Path to the Travco MSG file
    msg_file = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\Travel Agency TO\Travco\Hotel Booking Confirmation NU8B05A02.msg"
    
    print("Testing Travco MSG File Extraction")
    print("=" * 60)
    
    # Read the MSG file
    email_content = read_msg_file(msg_file)
    if not email_content:
        print("Failed to read MSG file")
        return
    
    # Test email identification
    subject = "Hotel Booking Confirmation NU8B05A02"
    sender = "travco@travco.co.uk"
    
    if is_travco_email(sender, subject):
        print("[OK] Email correctly identified as Travco email")
    else:
        print("[ERROR] Email NOT identified as Travco email")
    
    print("\nExtracting fields...")
    print("-" * 40)
    
    # Extract fields
    fields = extract_travco_fields(email_content, subject)
    
    # Display results
    for key, value in fields.items():
        if isinstance(value, float):
            print(f"{key}: AED {value:.2f}")
        else:
            print(f"{key}: {value}")
    
    print("\n" + "=" * 60)
    print("Extraction completed successfully!")

if __name__ == "__main__":
    main()