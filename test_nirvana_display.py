"""
Test script to display Nirvana parser results in the specified format
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Nirvana'))

from nirvana_parser import extract_nirvana_fields

def display_nirvana_results(file_path):
    """Display Nirvana parser results in specified format"""
    
    # Extract fields using Nirvana parser
    results = extract_nirvana_fields(file_path)
    
    # Display in the specified format
    print("NIRVANA BOOKING EXTRACTION RESULTS")
    print("=" * 50)
    print(f"MAIL_FIRST_NAME: {results.get('MAIL_FIRST_NAME', 'N/A')}")
    print(f"MAIL_FULL_NAME: {results.get('MAIL_FULL_NAME', 'N/A')}")
    print(f"MAIL_ARRIVAL: {results.get('MAIL_ARRIVAL', 'N/A')}")
    print(f"MAIL_DEPARTURE: {results.get('MAIL_DEPARTURE', 'N/A')}")
    print(f"MAIL_NIGHTS: {results.get('MAIL_NIGHTS', 'N/A')}")
    print(f"MAIL_PERSONS: {results.get('MAIL_PERSONS', 'N/A')}")
    print(f"MAIL_ROOM: {results.get('MAIL_ROOM', 'N/A')}")
    print(f"MAIL_RATE_CODE: {results.get('MAIL_RATE_CODE', 'N/A')}")
    print(f"MAIL_C_T_S: {results.get('MAIL_C_T_S', 'N/A')}")
    print(f"MAIL_NET_TOTAL: {results.get('MAIL_NET_TOTAL', 'N/A')}")
    print(f"MAIL_TOTAL: {results.get('MAIL_TOTAL', 'N/A')}")
    print(f"MAIL_TDF: {results.get('MAIL_TDF', 'N/A')}")
    print(f"MAIL_ADR: {results.get('MAIL_ADR', 'N/A')}")
    print(f"MAIL_AMOUNT: {results.get('MAIL_AMOUNT', 'N/A')}")

if __name__ == "__main__":
    # Test with the provided Nirvana booking file
    test_file = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Rules\Travel Agency TO\Nirvana\SB2508566475  Booking Confirmed.msg"
    
    print("Testing Nirvana Parser Integration")
    print("File:", test_file)
    print()
    
    display_nirvana_results(test_file)