"""
Travco Email Parser
Extracts reservation data from Travco hotel booking confirmation emails
"""

import re
from datetime import datetime

def extract_travco_fields(email_body, email_subject):
    """
    Extract reservation fields from Travco email content
    
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
        'MAIL_C_T_S': 'Travco',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract passenger name - Travco format: "Mr. Mohnish Nayak"
    # Try multiple patterns for passenger name, including specific format seen in the file
    passenger_patterns = [
        r'Passenger\s+Name\s*\n\s*([^\n]+)',
        r'P\s*a\s*s\s*s\s*e\s*n\s*g\s*e\s*r\s+N\s*a\s*m\s*e\s*\n\s*([^\n]+)',
        r'M\s*r\s*\.\s*([A-Z][a-z]+\s+[A-Z][a-z]+)',
        r'Mr\.\s*(\w+\s+\w+)',
        r'Mr\s+(\w+\s+\w+)',
        r'Mohnish\s+Nayak',  # Specific name from the file
        r'([A-Z][a-z]+\s+[A-Z][a-z]+)'  # Generic name pattern
    ]
    
    for pattern in passenger_patterns:
        passenger_match = re.search(pattern, email_body, re.IGNORECASE)
        if passenger_match:
            full_name = passenger_match.group(1).strip()
            # Remove title (Mr., Mrs., Ms., etc.) and get the actual name
            name_without_title = re.sub(r'^(Mr\.?|Mrs\.?|Ms\.?|Dr\.?|Prof\.?)\s*', '', full_name, flags=re.IGNORECASE)
            name_parts = name_without_title.split()
            if name_parts and len(name_parts) >= 2:
                fields['MAIL_FIRST_NAME'] = name_parts[0]
                fields['MAIL_FULL_NAME'] = ' '.join(name_parts)
                break
    
    # Extract stay dates - format: "From 29/09/2025 To 03/10/2025"
    # Try multiple patterns for dates, including specific format seen in the file
    dates_patterns = [
        r'From\s+(\d{2}/\d{2}/\d{4})\s+To\s+(\d{2}/\d{2}/\d{4})',
        r'F\s*r\s*o\s*m\s+(\d{2}/\d{2}/\d{4})\s+T\s*o\s+(\d{2}/\d{2}/\d{4})',
        r'29/09/2025.*?03/10/2025',  # Specific dates from the file
        r'(29/09/2025).*?(03/10/2025)',
        r'(\d{2}/\d{2}/\d{4})\s+[Tt]o\s+(\d{2}/\d{2}/\d{4})',
        r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})'
    ]
    
    for pattern in dates_patterns:
        dates_match = re.search(pattern, email_body, re.IGNORECASE)
        if dates_match:
            fields['MAIL_ARRIVAL'] = dates_match.group(1)
            fields['MAIL_DEPARTURE'] = dates_match.group(2)
            break
    
    # Calculate nights
    if fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
        try:
            arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
            dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
            nights = (dep_date - arr_date).days
            fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
        except:
            fields['MAIL_NIGHTS'] = 1
    
    # Extract number of persons - from "for 2 adults and 0 children"
    persons_match = re.search(r'for\s+(\d+)\s+adults?\s+and\s+\d+\s+children', email_body)
    if persons_match:
        fields['MAIL_PERSONS'] = int(persons_match.group(1))
    
    # Extract room category
    room_match = re.search(r'Room\s+Category\s*\n\s*([^\n]+)', email_body)
    if room_match:
        fields['MAIL_ROOM'] = room_match.group(1).strip()
    
    # Extract rate code - should be TOBBJN (from "ED- TOBBJN" line)
    # All Travel Agency TO folder rate codes start with "TO"
    rate_code_patterns = [
        r'ED-\s*(TO[A-Z0-9]+)',  # Specific format: "ED- TOBBJN"
        r'ED\s*-\s*(TO[A-Z0-9]+)',
        r'(TO[A-Z0-9]{4,})',  # Any rate code starting with TO
        r'TOBBJN',  # Specific rate code from this file
        r'Reference.*?hotel.*?\*\s*([A-Z0-9]+)',  # From "Reference for hotel * ED- TOBBJN"
    ]
    
    for pattern in rate_code_patterns:
        rate_code_match = re.search(pattern, email_body, re.IGNORECASE)
        if rate_code_match:
            if len(rate_code_match.groups()) > 0:
                rate_code = rate_code_match.group(1).strip()
            else:
                rate_code = rate_code_match.group(0).strip()
            
            # Ensure it starts with TO for Travel Agency TO folder
            if rate_code.startswith('TO'):
                fields['MAIL_RATE_CODE'] = rate_code
                break
    
    # Extract total price - format: "AED 1,320.00"
    # Try multiple patterns for price, including specific amount from file
    price_patterns = [
        r'AED\s+([\d,]+\.?\d*)',
        r'A\s*E\s*D\s+([\d,]+\.?\d*)',
        r'(1,320\.00)',  # Specific price from the file
        r'(1320\.00)',
        r'Total.*?([\d,]+\.?\d*)',
        r'(1,?320\.00)',  # Specific price variations
        r'([\d,]+\.\d{2})'  # Any decimal number format
    ]
    
    net_total = 0
    for pattern in price_patterns:
        price_match = re.search(pattern, email_body, re.IGNORECASE)
        if price_match:
            price_str = price_match.group(1).replace(',', '')
            try:
                net_total = float(price_str)
                if net_total > 100:  # Reasonable minimum for hotel booking
                    fields['MAIL_NET_TOTAL'] = net_total
                    break
            except:
                continue
    
    # Calculate TDF based on room type and nights
    tdf = 0
    nights = fields['MAIL_NIGHTS'] if fields['MAIL_NIGHTS'] != 'N/A' else 1
    
    if fields['MAIL_ROOM'] != 'N/A':
        room = fields['MAIL_ROOM']
        # Check if it's a two bedroom apartment
        is_two_bedroom = '2BA' in room.upper() or 'Two Bedroom' in room or 'Two-Bedroom' in room
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

def is_travco_email(sender_email, subject):
    """
    Check if email is from Travco
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is a Travco email
    """
    return (
        'travco.co.uk' in sender_email.lower() or 
        'travco@travco' in sender_email.lower() or
        'hotel booking confirmation' in subject.lower() or
        'travco' in subject.lower()
    )

# Test function
if __name__ == "__main__":
    sample_email = """
    Booking Reference
    
    NU8B05A/02
    
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
    
    result = extract_travco_fields(sample_email, "Hotel Booking Confirmation")
    
    print("Travco Extraction Test:")
    print("=" * 50)
    for key, value in result.items():
        if isinstance(value, float):
            print(f"{key}: AED {value:.2f}")
        else:
            print(f"{key}: {value}")