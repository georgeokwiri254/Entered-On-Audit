"""
Duri Travel (Dakkak DMC) Email Parser
Extracts reservation data from Duri Travel booking confirmation emails from Dakkak DMC
Based on analysis of Hotel New Booking format
"""

import re
from datetime import datetime

def extract_dakkak_fields(email_body, email_subject):
    """
    Extract reservation fields from Duri Travel email content
    
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
        'MAIL_C_T_S': 'Duri Travel',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract guest name - Format: "Mr.SAVINI ENRICO" from FILE NO section
    # or from "Guest Name 1 Mr.SAVINI ENRICO - (Adult)"
    name_patterns = [
        r'Guest Name 1\s+([^-]+)',  # Guest Name 1 Mr.SAVINI ENRICO - (Adult)
        r'Mr\.([A-Z\s]+)(?:\s+-\s+\(Adult\))?',  # Mr.SAVINI ENRICO - (Adult)
        r'([A-Z]+\s+[A-Z]+)\s+ITALY',  # SAVINI ENRICO ITALY from passport section
        r'Mr\.?\s*([A-Z][a-z]+\s+[A-Z][a-z]+)',  # Mr. FirstName LastName
        r'([A-Z][A-Z\s]+)',  # All caps names
    ]
    
    for pattern in name_patterns:
        name_match = re.search(pattern, email_body)
        if name_match:
            full_name = name_match.group(1).strip()
            # Remove "Mr." prefix if present
            full_name = re.sub(r'^Mr\.?\s*', '', full_name).strip()
            
            # Clean up the name and split
            name_parts = full_name.split()
            if len(name_parts) >= 2:
                # For names like "SAVINI ENRICO", first word is surname, second is first name
                fields['MAIL_FIRST_NAME'] = name_parts[-1]  # Last word as first name (ENRICO)
                fields['MAIL_FULL_NAME'] = full_name
                break
    
    # Extract dates - Format: "07-Nov-2025" to "11-Nov-2025"
    date_patterns = [
        r'(\d{2}-[A-Z][a-z]{2}-\d{4})\s+(\d{2}-[A-Z][a-z]{2}-\d{4})',  # 07-Nov-2025 11-Nov-2025
        r'From Date.*?(\d{2}-[A-Z][a-z]{2}-\d{4}).*?To Date.*?(\d{2}-[A-Z][a-z]{2}-\d{4})',
    ]
    
    for pattern in date_patterns:
        dates_match = re.search(pattern, email_body, re.DOTALL)
        if dates_match:
            arrival_str = dates_match.group(1)
            departure_str = dates_match.group(2)
            
            # Convert date format from "07-Nov-2025" to "07/11/2025"
            try:
                arrival_date = datetime.strptime(arrival_str, '%d-%b-%Y')
                departure_date = datetime.strptime(departure_str, '%d-%b-%Y')
                
                fields['MAIL_ARRIVAL'] = arrival_date.strftime('%d/%m/%Y')
                fields['MAIL_DEPARTURE'] = departure_date.strftime('%d/%m/%Y')
                break
            except:
                continue
    
    # Extract nights - directly from the nights column or calculate
    nights_match = re.search(r'Night\(s\)\s+.*?\s+(\d+)\s+', email_body)
    if nights_match:
        fields['MAIL_NIGHTS'] = int(nights_match.group(1))
    elif fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
        try:
            arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
            dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
            nights = (dep_date - arr_date).days
            fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
        except:
            fields['MAIL_NIGHTS'] = 1
    
    # Extract persons - from "Total Adults" column or from room description
    persons_patterns = [
        r'Total Adults\s+.*?\s+(\d+)\s+',  # Total Adults column
        r'(\d+)\s+Adult',  # "2 Adult + 0 Child"
        r'(\d+)\s+adults?',  # Generic adult pattern
    ]
    
    for pattern in persons_patterns:
        persons_match = re.search(pattern, email_body)
        if persons_match:
            fields['MAIL_PERSONS'] = int(persons_match.group(1))
            break
    
    # Extract room type - from "Room Catg." column
    room_patterns = [
        r'Room Catg\.\s+.*?\s+(Superior Room[^|]+)',
        r'(Superior Room-Breakfast Included)',
        r'Superior Room-Breakfast Included',
    ]
    
    for pattern in room_patterns:
        room_match = re.search(pattern, email_body)
        if room_match:
            if len(room_match.groups()) > 0:
                fields['MAIL_ROOM'] = room_match.group(1).strip()
            else:
                fields['MAIL_ROOM'] = room_match.group(0).strip()
            break
    
    # Extract rate code - prioritize BKGHO file number over promo code
    file_match = re.search(r'(BKGHO\d+)', email_body)
    if file_match:
        fields['MAIL_RATE_CODE'] = file_match.group(1)
    else:
        # Extract promo code - from "Promo Code:" line (may be empty)
        promo_match = re.search(r'Promo Code:\s*([^\s]+)', email_body)
        if promo_match and promo_match.group(1).strip() and promo_match.group(1).strip() not in ['', '-']:
            fields['MAIL_RATE_CODE'] = promo_match.group(1).strip()
    
    # Extract total price - Format: "AED 3600.00"
    price_patterns = [
        r'Total Amount\s*:\s*AED\s+([\d,]+\.?\d*)',
        r'Total Price.*?AED\s+([\d,]+\.?\d*)',
        r'AED\s+([\d,]+\.?\d*)',
    ]
    
    net_total = 0
    for pattern in price_patterns:
        price_match = re.search(pattern, email_body)
        if price_match:
            price_str = price_match.group(1).replace(',', '')
            try:
                net_total = float(price_str)
                if net_total > 100:  # Reasonable minimum for hotel booking
                    fields['MAIL_NET_TOTAL'] = net_total
                    break
            except:
                continue
    
    # Calculate TDF based on room type and nights (same logic as other parsers)
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
    
    # Calculate derived values (same logic as other parsers)
    if net_total > 0:
        mail_total = net_total + tdf
        mail_amount = net_total / 1.225  # Remove VAT
        mail_adr = mail_amount / nights if nights > 0 else 0
        
        fields['MAIL_TOTAL'] = mail_total
        fields['MAIL_AMOUNT'] = mail_amount
        fields['MAIL_ADR'] = mail_adr
    
    return fields

def is_duri_travel_email(sender_email, subject):
    """
    Check if email is from Duri Travel / Dakkak DMC
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is a Duri Travel email
    """
    return (
        'dakkak' in sender_email.lower() or 
        'duritravel' in sender_email.lower() or
        'dakkak dmc' in subject.lower() or
        'hotel new booking' in subject.lower() and 'dakkak' in subject.lower() or
        'BKGHO' in subject  # Booking reference format
    )

# Test function
if __name__ == "__main__":
    # Sample data extracted from the actual Duri Travel email
    sample_email = """
    HOTEL BOOKING 
    
    ( Grand Millennium Dubai )
    
    ( Vouchered )
    
    FILE NO     ISSUED DATE     Guest Name     Pax passport     
    BKGHO0825397     2025-08-25     Mr.SAVINI ENRICO     ITALY     
    
    Room 1     
    From Date     To Date     Night(s)     Room Catg.     Meal Type     Total Adults     Total Child     Use Allotment     Total Price     
    07-Nov-2025     11-Nov-2025     4     Superior Room-Breakfast Included     Breakfast Included     2     0     Y     AED 3600.00     
    Room 1: Superior Room-Breakfast Included ( 2 Adult + 0 Child )     Promo Code: 
    
    Guest Name :     Guest Name 1 Mr.SAVINI ENRICO - (Adult)
    Guest Name 2 Mrs.MAMI CAROL - (Adult)
    
    Fri     07-Nov-2025     AED 900.00     
    Sat     08-Nov-2025     AED 900.00     
    Sun     09-Nov-2025     AED 900.00     
    Mon     10-Nov-2025     AED 900.00     
    
    Total Amount : AED 3600.00     
    
    Remark : -     
    SpecialRequest : -     
    """
    
    result = extract_duri_travel_fields(sample_email, "Hotel New Booking Mr.SAVINI ENRICO BKGHO0825397 - (Dakkak DMC)")
    
    print("Duri Travel Extraction Test:")
    print("=" * 50)
    
    # Display in the requested format
    print(f"MAIL_FIRST_NAME: {result['MAIL_FIRST_NAME']}")
    print(f"MAIL_FULL_NAME: {result['MAIL_FULL_NAME']}")
    print(f"MAIL_ARRIVAL: {result['MAIL_ARRIVAL']}")
    print(f"MAIL_DEPARTURE: {result['MAIL_DEPARTURE']}")
    print(f"MAIL_NIGHTS: {result['MAIL_NIGHTS']}")
    print(f"MAIL_PERSONS: {result['MAIL_PERSONS']}")
    print(f"MAIL_ROOM: {result['MAIL_ROOM']}")
    print(f"MAIL_RATE_CODE: {result['MAIL_RATE_CODE']}")
    print(f"MAIL_C_T_S: {result['MAIL_C_T_S']}")
    
    # Format currency values
    net_total = result['MAIL_NET_TOTAL']
    total = result['MAIL_TOTAL']
    tdf = result['MAIL_TDF']
    adr = result['MAIL_ADR']
    amount = result['MAIL_AMOUNT']
    
    print(f"MAIL_NET_TOTAL: {f'AED {net_total:.2f}' if isinstance(net_total, (int, float)) else net_total}")
    print(f"MAIL_TOTAL: {f'AED {total:.2f}' if isinstance(total, (int, float)) else total}")
    print(f"MAIL_TDF: {f'AED {tdf:.2f}' if isinstance(tdf, (int, float)) else tdf}")
    print(f"MAIL_ADR: {f'AED {adr:.2f}' if isinstance(adr, (int, float)) else adr}")
    print(f"MAIL_AMOUNT: {f'AED {amount:.2f}' if isinstance(amount, (int, float)) else amount}")