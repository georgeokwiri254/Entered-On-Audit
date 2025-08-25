"""
Travco Manual Extraction Template
This provides a template for manual extraction when MSG parsing is difficult
"""

def extract_travco_manual_template():
    """
    Template for manual Travco booking extraction
    Fill in the values based on the booking confirmation email
    """
    
    # Known values from Hotel Booking Confirmation NU8B05A02.msg
    travco_booking_data = {
        'MAIL_FIRST_NAME': 'Mohnish',
        'MAIL_FULL_NAME': 'Mohnish Nayak',
        'MAIL_ARRIVAL': '29/09/2025',
        'MAIL_DEPARTURE': '03/10/2025',
        'MAIL_NIGHTS': 4,  # Calculated: (03/10/2025 - 29/09/2025)
        'MAIL_PERSONS': 2,  # From "for 2 adults and 0 children"
        'MAIL_ROOM': 'Superior - Bed & Breakfast - Double',
        'MAIL_RATE_CODE': 'TOBBJN',  # From "ED- TOBBJN" line
        'MAIL_C_T_S': 'Travco',
        'MAIL_NET_TOTAL': 1320.00,  # AED 1,320.00
        'MAIL_TDF': 80,  # 4 nights × 20 AED (standard room)
        'MAIL_TOTAL': 1400.00,  # NET_TOTAL + TDF
        'MAIL_AMOUNT': 1077.55,  # NET_TOTAL / 1.225 (excluding VAT)
        'MAIL_ADR': 269.39  # MAIL_AMOUNT / nights
    }
    
    return travco_booking_data

def display_extraction_format():
    """Display the required format for Travco extractions"""
    
    print("Travco Email Extraction Format:")
    print("=" * 50)
    
    data = extract_travco_manual_template()
    
    for key, value in data.items():
        if isinstance(value, float):
            print(f"{key}: AED {value:.2f}")
        else:
            print(f"{key}: {value}")
    
    print("\n" + "=" * 50)
    print("Extraction Template for Future Travco Emails:")
    print("=" * 50)
    
    template = {
        'MAIL_FIRST_NAME': '[Extract first name from Passenger Name]',
        'MAIL_FULL_NAME': '[Extract full name from Passenger Name]',
        'MAIL_ARRIVAL': '[Extract from Stay Dates - From date]',
        'MAIL_DEPARTURE': '[Extract from Stay Dates - To date]',
        'MAIL_NIGHTS': '[Calculate: departure - arrival]',
        'MAIL_PERSONS': '[Extract from "for X adults"]',
        'MAIL_ROOM': '[Extract from Room Category]',
        'MAIL_RATE_CODE': '[Extract from "ED- TOBBJN" line - should start with TO]',
        'MAIL_C_T_S': 'Travco',
        'MAIL_NET_TOTAL': '[Extract total price amount]',
        'MAIL_TDF': '[Calculate: nights × 20 (or 40 for 2BR)]',
        'MAIL_TOTAL': '[Calculate: NET_TOTAL + TDF]',
        'MAIL_AMOUNT': '[Calculate: NET_TOTAL / 1.225]',
        'MAIL_ADR': '[Calculate: MAIL_AMOUNT / nights]'
    }
    
    for key, value in template.items():
        print(f"{key}: {value}")

if __name__ == "__main__":
    display_extraction_format()