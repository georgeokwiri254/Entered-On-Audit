"""
Traveltino Email Parser
Extracts reservation data from Traveltino confirmation emails
"""

import re
from datetime import datetime

def extract_traveltino_fields(email_body, email_subject):
    """
    Extract reservation fields from Traveltino email content
    
    Args:
        email_body (str): Email body content
        email_subject (str): Email subject
    
    Returns:
        dict: Extracted field values
    """
    
    # Initialize result dictionary
    fields = {
        'MAIL_FIRST_NAME': 'N/A',
        'MAIL_FULL_NAME': 'N/A', 
        'MAIL_ARRIVAL': 'N/A',
        'MAIL_DEPARTURE': 'N/A',
        'MAIL_NIGHTS': 'N/A',
        'MAIL_PERSONS': 'N/A',
        'MAIL_ROOM': 'N/A',
        'MAIL_RATE_CODE': 'N/A',
        'MAIL_C_T_S': 'Traveltino',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract names - Traveltino specific mapping
    first_name_match = re.search(r'Name:\s*([A-Z\s]+)', email_body)
    last_name_match = re.search(r'Last Name:\s*([A-Z\s]+)', email_body)
    
    # For Traveltino: MAIL_FIRST_NAME = Name field, MAIL_FULL_NAME = Last Name field
    if first_name_match:
        fields['MAIL_FIRST_NAME'] = first_name_match.group(1).strip()
    if last_name_match:
        fields['MAIL_FULL_NAME'] = last_name_match.group(1).strip()
    
    # Extract dates
    arrival_match = re.search(r'Arrival Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    departure_match = re.search(r'Departure Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    
    if arrival_match:
        fields['MAIL_ARRIVAL'] = arrival_match.group(1)
    if departure_match:
        fields['MAIL_DEPARTURE'] = departure_match.group(1)
    
    # Calculate nights
    if fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
        try:
            arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
            dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
            nights = (dep_date - arr_date).days
            fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
        except:
            fields['MAIL_NIGHTS'] = 1
    
    # Extract persons
    persons_match = re.search(r'\((\d+) Adult\)', email_body)
    if persons_match:
        fields['MAIL_PERSONS'] = int(persons_match.group(1))
    
    # Extract room type
    room_match = re.search(r'(\d+ x [^(]+\([^)]+\)[^)]*)', email_body)
    if room_match:
        fields['MAIL_ROOM'] = room_match.group(1).strip()
    
    # Extract promo code (rate code)
    promo_match = re.search(r'Promo code:\s*([A-Z0-9{}\s]+)', email_body)
    if promo_match:
        fields['MAIL_RATE_CODE'] = promo_match.group(1).strip()
    
    # Extract booking cost (net total)
    cost_match = re.search(r'Booking cost price:\s*([\d,.]+)\s*AED', email_body)
    net_total = 0
    if cost_match:
        net_total = float(cost_match.group(1).replace(',', ''))
        fields['MAIL_NET_TOTAL'] = net_total
    
    # Calculate TDF based on room type and nights
    tdf = 0
    nights = fields['MAIL_NIGHTS'] if fields['MAIL_NIGHTS'] != 'N/A' else 1
    
    if fields['MAIL_ROOM'] != 'N/A':
        room = fields['MAIL_ROOM']
        is_two_bedroom = '2BA' in room.upper() or 'Two Bedroom' in room
        tdf_rate = 40 if is_two_bedroom else 20
        
        # For 30+ nights, use 30 as the multiplier instead of actual nights
        effective_nights = min(nights, 30) if nights >= 30 else nights
        tdf = effective_nights * tdf_rate
        fields['MAIL_TDF'] = tdf
    
    # Calculate derived values
    if net_total > 0:
        mail_total = net_total + tdf
        mail_amount = net_total / 1.225
        mail_adr = mail_amount / nights if nights > 0 else 0
        
        fields['MAIL_TOTAL'] = mail_total
        fields['MAIL_AMOUNT'] = mail_amount
        fields['MAIL_ADR'] = mail_adr
    
    return fields

def is_traveltino_email(sender_email, subject):
    """
    Check if email is from Traveltino
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is a Traveltino email
    """
    return (
        'traveltino' in sender_email.lower() or 
        'traveltino' in subject.lower() or
        'booking confirmation' in subject.lower()
    )

# Test function
if __name__ == "__main__":
    sample_email = """
    Name: CUSTOMER
    Last Name: NAME
    
    Arrival Date: 27/08/2025
    Departure Date: 28/08/2025
    Rooms: 1 x Superior Room (King/Twin) - Double (1 Adult)
    
    Booking cost price: 200.00 AED
    
    Promo code: TOBBJN{ALL MARKET EX UAE}
    """
    
    result = extract_traveltino_fields(sample_email, "1055683 - Grand Millennium Dubai booking confirmation")
    
    print("Traveltino Extraction Test:")
    print("=" * 50)
    for key, value in result.items():
        if isinstance(value, float):
            print(f"{key}: AED {value:.2f}")
        else:
            print(f"{key}: {value}")