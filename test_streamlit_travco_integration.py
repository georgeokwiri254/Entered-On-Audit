"""
Test script to verify Travco parser integration in Streamlit app
"""

import sys
from pathlib import Path

# Add the current directory to path
sys.path.insert(0, str(Path(__file__).parent))

# Import the extraction function from streamlit app
from streamlit_app import extract_reservation_fields

def test_travco_integration():
    """Test Travco parser integration"""
    print("Testing Travco Parser Integration")
    print("=" * 50)
    
    # Test email content (based on the actual MSG file)
    sample_email_body = """
    CAUTION: This is an external email from travco@travco.co.uk
    
    Booking Reference
    
    NU8B05A/02
    
    Reference for hotel
    
    ED- TOBBJN
    
    Passenger Name
    
    Mr. Mohnish Nayak
    
    Hotel Name
    
    Grand Millennium Dubai
    
    Room Category
    
    Superior - Bed & Breakfast - Double
    
    Number of rooms
    
    1
    
    Stay Dates
    
    From 29/09/2025 To 03/10/2025
    
    Total Price of booking (including any discounts) for 2 adults and 0 children
    
    AED 1,320.00
    """
    
    # Test sender email
    sender_email = "travco@travco.co.uk"
    
    print("Test Input:")
    print(f"Sender: {sender_email}")
    print(f"Body length: {len(sample_email_body)} characters")
    print()
    
    # Extract fields
    result = extract_reservation_fields(sample_email_body, sender_email)
    
    print("Extraction Results:")
    print("-" * 30)
    for key, value in result.items():
        print(f"{key}: {value}")
    
    print("\n" + "=" * 50)
    
    # Verify key fields are extracted correctly
    expected_fields = {
        'C_T_S': 'Travco',
        'RATE_CODE': 'TOBBJN',
        'FIRST_NAME': 'Mohnish',
        'FULL_NAME': 'Mohnish Nayak',
        'ARRIVAL': '29/09/2025',
        'DEPARTURE': '03/10/2025',
        'NIGHTS': 4,
        'PERSONS': 2,
        'ROOM': 'Superior - Bed & Breakfast - Double',
        'NET_TOTAL': 1320.00
    }
    
    print("Verification:")
    print("-" * 30)
    all_passed = True
    
    for field, expected in expected_fields.items():
        actual = result.get(field, 'NOT_FOUND')
        if actual == expected:
            print(f"[PASS] {field}: {actual}")
        else:
            print(f"[FAIL] {field}: Expected {expected}, Got {actual}")
            all_passed = False
    
    print("\n" + "=" * 50)
    if all_passed:
        print("SUCCESS: ALL TESTS PASSED! Travco integration working correctly.")
    else:
        print("WARNING: Some tests failed. Check the extraction logic.")
    
    return result

def test_dubai_link_integration():
    """Test Dubai Link parser integration"""
    print("\nTesting Dubai Link Parser Integration")
    print("=" * 50)
    
    # Test Dubai Link email content
    sample_email_body = """
    Name: SOHEIL
    Last Name: RADIOM
    
    Arrival Date: 27/08/2025
    Departure Date: 28/08/2025
    Rooms: 1 x Superior Room (King/Twin) - Double (1 Adult)
    
    Booking cost price: 200.00 AED
    
    Promo code: TOBBJN{ALL MARKET EX UAE}
    """
    
    # Test sender email
    sender_email = "reservations@gte.travel"
    
    print("Test Input:")
    print(f"Sender: {sender_email}")
    print(f"Body length: {len(sample_email_body)} characters")
    print()
    
    # Extract fields
    result = extract_reservation_fields(sample_email_body, sender_email)
    
    print("Extraction Results:")
    print("-" * 30)
    for key, value in result.items():
        if value != 'N/A':
            print(f"{key}: {value}")
    
    return result

if __name__ == "__main__":
    # Test Travco integration
    travco_result = test_travco_integration()
    
    # Test Dubai Link integration
    dubai_result = test_dubai_link_integration()