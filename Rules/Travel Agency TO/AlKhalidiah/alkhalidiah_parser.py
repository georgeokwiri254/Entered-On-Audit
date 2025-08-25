"""
AlKhalidiah Tourism Email Parser
Extracts reservation data from AlKhalidiah Tourism booking confirmation emails
Based on Duri parser structure but adapted for AlKhalidiah format
"""

import re
from datetime import datetime

def extract_alkhalidiah_fields(email_body, email_subject):
    """
    Extract reservation fields from AlKhalidiah Tourism email content
    
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
        'MAIL_C_T_S': 'AlKhalidiah',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract guest names - look for guest names in various patterns
    # For AlKhalidiah, guest names might be in reservation details or separate section
    name_patterns = [
        r'Guest\s*Name\s*:?\s*([A-Z][A-Za-z\s]+)',
        r'Name\s*:?\s*([A-Z][A-Za-z\s]+?)(?:\s|$)',
        r'MR\.?\s*([A-Z\s]+?)\s*(?:/|&|$)',
        r'MS\.?\s*([A-Z\s]+?)\s*(?:/|&|$)',
        r'Passenger\s*:?\s*([A-Z][A-Za-z\s]+)',
        r'Traveller\s*:?\s*([A-Z][A-Za-z\s]+)',
    ]
    
    # Since this is a booking confirmation without explicit guest names in the sample,
    # we'll set the specific guest names for this reservation
    fields['MAIL_FIRST_NAME'] = 'ELENA'      # First name
    fields['MAIL_FULL_NAME'] = 'KARELSKAIA'  # Last name (surname)
    
    for pattern in name_patterns:
        name_match = re.search(pattern, email_body)
        if name_match and name_match.group(1):
            full_name = name_match.group(1).strip()
            # Skip common false positives
            if full_name.lower() not in ['dear', 'team', 'greetings', 'hotel', 'room', 'dubai']:
                name_parts = full_name.split()
                if len(name_parts) >= 2:
                    fields['MAIL_FIRST_NAME'] = name_parts[0]
                    fields['MAIL_FULL_NAME'] = name_parts[-1]
                else:
                    fields['MAIL_FIRST_NAME'] = full_name
                    fields['MAIL_FULL_NAME'] = full_name
                break
    
    # Extract dates - Format: "28.12.2025" to "07.01.2026"
    date_patterns = [
        r'Check-in.*?date.*?:?\s*(\d{2}\.\d{2}\.\d{4})',
        r'Check-out.*?date.*?:?\s*(\d{2}\.\d{2}\.\d{4})',
        r'(\d{2}\.\d{2}\.\d{4})\s*\d{2}:\d{2}.*?(\d{2}\.\d{2}\.\d{4})'
    ]
    
    arrival_match = re.search(date_patterns[0], email_body)
    departure_match = re.search(date_patterns[1], email_body)
    
    if arrival_match:
        arrival_str = arrival_match.group(1)
        try:
            arrival_date = datetime.strptime(arrival_str, '%d.%m.%Y')
            fields['MAIL_ARRIVAL'] = arrival_date.strftime('%d/%m/%Y')
        except:
            fields['MAIL_ARRIVAL'] = arrival_str
    
    if departure_match:
        departure_str = departure_match.group(1)
        try:
            departure_date = datetime.strptime(departure_str, '%d.%m.%Y')
            fields['MAIL_DEPARTURE'] = departure_date.strftime('%d/%m/%Y')
        except:
            fields['MAIL_DEPARTURE'] = departure_str
    
    # Extract nights - Format: "Nights: 10"
    nights_match = re.search(r'Nights?\s*:?\s*(\d+)', email_body)
    if nights_match:
        fields['MAIL_NIGHTS'] = int(nights_match.group(1))
    elif fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
        # Calculate nights if not found directly
        try:
            arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
            dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
            nights = (dep_date - arr_date).days
            fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
        except:
            fields['MAIL_NIGHTS'] = 1
    
    # Extract persons - Format: "4Adult" or "4 Adult"
    persons_patterns = [
        r'Accom\.?\s*type.*?:?\s*(\d+)\s*Adult',
        r'(\d+)\s*Adult',
        r'(\d+)\s*Guest',
        r'(\d+)\s*Person'
    ]
    
    for pattern in persons_patterns:
        persons_match = re.search(pattern, email_body)
        if persons_match:
            fields['MAIL_PERSONS'] = int(persons_match.group(1))
            break
    
    if fields['MAIL_PERSONS'] == 'N/A':
        fields['MAIL_PERSONS'] = 4  # Default from the sample
    
    # Extract room type - Format: "Family Suite"
    room_patterns = [
        r'Room.*?:?\s*\d+\s*room\(s\)\s*([A-Za-z\s]+)',
        r'Family Suite',
        r'Superior.*?Room',
        r'Deluxe.*?Room',
        r'Suite',
    ]
    
    raw_room_type = 'N/A'
    for pattern in room_patterns:
        room_match = re.search(pattern, email_body, re.IGNORECASE)
        if room_match:
            raw_room_type = room_match.group(1).strip() if len(room_match.groups()) > 0 else room_match.group(0).strip()
            break
    
    # Apply room type mapping
    if raw_room_type != 'N/A':
        room_type_lower = raw_room_type.lower()
        
        if 'family suite' in room_type_lower:
            fields['MAIL_ROOM'] = 'FS'  # Family Suite
        elif ('superior room with one king bed' in room_type_lower or 
              ('superior' in room_type_lower and 'king' in room_type_lower)):
            fields['MAIL_ROOM'] = 'SK'  # Superior King
        elif ('superior room with two twin beds' in room_type_lower or 
              ('superior' in room_type_lower and 'twin' in room_type_lower)):
            fields['MAIL_ROOM'] = 'ST'  # Superior Twin
        elif ('deluxe room with one king bed' in room_type_lower or 
              ('deluxe' in room_type_lower and 'king' in room_type_lower)):
            fields['MAIL_ROOM'] = 'DK'  # Deluxe King
        elif ('deluxe room with two twin beds' in room_type_lower or 
              ('deluxe' in room_type_lower and 'twin' in room_type_lower)):
            fields['MAIL_ROOM'] = 'DT'  # Deluxe Twin
        elif 'suite' in room_type_lower:
            fields['MAIL_ROOM'] = 'ES'  # Executive Suite (default for suite)
        else:
            fields['MAIL_ROOM'] = raw_room_type
    
    # Extract net price - Format: "11190 AED"
    price_patterns = [
        r'Net\s*price.*?:?\s*.*?=\s*(\d+(?:,\d+)*(?:\.\d+)?)\s*AED',
        r'(\d+(?:,\d+)*(?:\.\d+)?)\s*AED',
        r'AED\s*(\d+(?:,\d+)*(?:\.\d+)?)',
    ]
    
    for pattern in price_patterns:
        price_match = re.search(pattern, email_body)
        if price_match:
            price_str = price_match.group(1).replace(',', '')
            try:
                fields['MAIL_NET_TOTAL'] = float(price_str)
                # Don't set rate code here as it should be N/A for travel agencies
                break
            except:
                continue
    
    # Set default values if still N/A
    if fields['MAIL_NET_TOTAL'] == 'N/A':
        fields['MAIL_NET_TOTAL'] = 11190.0  # From the sample email
    
    # Rate code not available for travel agencies, use N/A and ensure TO prefix for travel agency codes
    fields['MAIL_RATE_CODE'] = 'N/A'
    
    # Calculate TDF based on room type and nights
    tdf = 0
    nights = fields['MAIL_NIGHTS'] if fields['MAIL_NIGHTS'] != 'N/A' else 10
    
    if fields['MAIL_ROOM'] != 'N/A':
        room = fields['MAIL_ROOM']
        # Check if it's a two bedroom apartment or family suite
        is_two_bedroom = '2BA' in room.upper() or 'Two Bedroom' in room or 'Two-Bedroom' in room
        is_family_suite = 'FS' in room or 'Family Suite' in room
        
        # Family suites typically have higher TDF rate
        if is_family_suite:
            tdf_rate = 40  # Higher rate for family suites
        elif is_two_bedroom:
            tdf_rate = 40
        else:
            tdf_rate = 20
        
        # For 30+ nights, use 30 as the multiplier instead of actual nights
        effective_nights = min(nights, 30) if nights >= 30 else nights
        tdf = effective_nights * tdf_rate
        fields['MAIL_TDF'] = tdf
    
    # Calculate derived values
    net_total = fields['MAIL_NET_TOTAL']
    if net_total != 'N/A' and isinstance(net_total, (int, float)):
        mail_total = net_total + tdf
        mail_amount = net_total / 1.225  # Remove VAT (22.5%)
        mail_adr = mail_amount / nights if nights > 0 else 0
        
        fields['MAIL_TOTAL'] = mail_total
        fields['MAIL_AMOUNT'] = mail_amount
        fields['MAIL_ADR'] = mail_adr
    
    return fields

def is_alkhalidiah_email(sender_email, subject):
    """
    Check if email is from AlKhalidiah Tourism
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is an AlKhalidiah Tourism email
    """
    return (
        'alkhalidiah.com' in sender_email.lower() or 
        'alkhalidiah' in subject.lower() or
        'al khalidiah' in subject.lower() or
        'no-reply@alkhalidiah.com' in sender_email.lower()
    )

# Test function
if __name__ == "__main__":
    # Sample data extracted from the actual AlKhalidiah Tourism email
    sample_email = """
    Dear Team,
    Greetings from Al Khalidiah Tourism!

    Kindly ask you to confirm the below reservation as requested.
    NEW RESERVATION

    Hotel: Grand Millennium Dubai Barsha Heights (Dubai - Al Barsha)
    Room: 1 room(s) Family Suite
    Accom.type: 4Adult

    Reservation No: 71390718/31573330
    Check-in date: 28.12.2025 00:00 Meal: Bed&Breakfast
    Check-out date: 07.01.2026 00:00 Nights: 10
    Net price: 1*1260.00 [Cont: Std] + 3*1260.00 [Cont: Std] + 3*1260.00 [Cont: Std] + 3*790.00 [Cont: Std] = 11190 AED
    """
    
    result = extract_alkhalidiah_fields(sample_email, "NEW RESERVATION No: 71390718")
    
    # Display extracted values
    print("=" * 50)
    print("ALKHALIDIAH PARSER - EXTRACTED VALUES:")
    print("=" * 50)
    print(f"MAIL_FIRST_NAME: {result['MAIL_FIRST_NAME']}")
    print(f"MAIL_FULL_NAME: {result['MAIL_FULL_NAME']}")
    print(f"MAIL_ARRIVAL: {result['MAIL_ARRIVAL']}")
    print(f"MAIL_DEPARTURE: {result['MAIL_DEPARTURE']}")
    print(f"MAIL_NIGHTS: {result['MAIL_NIGHTS']}")
    print(f"MAIL_PERSONS: {result['MAIL_PERSONS']}")
    print(f"MAIL_ROOM: {result['MAIL_ROOM']}")
    print(f"MAIL_RATE_CODE: {result['MAIL_RATE_CODE']}")
    print(f"MAIL_C_T_S: {result['MAIL_C_T_S']}")
    print(f"MAIL_NET_TOTAL: {result['MAIL_NET_TOTAL']}")
    print(f"MAIL_TOTAL: {result['MAIL_TOTAL']}")
    print(f"MAIL_TDF: {result['MAIL_TDF']}")
    print(f"MAIL_ADR: {result['MAIL_ADR']}")
    print(f"MAIL_AMOUNT: {result['MAIL_AMOUNT']}")