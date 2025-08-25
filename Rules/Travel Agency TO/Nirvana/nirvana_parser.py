"""
Nirvana Email Parser
Extracts reservation data from Nirvana booking confirmation emails
"""

import re
import extract_msg
import os
from datetime import datetime

def convert_nirvana_date(date_str):
    """Convert date format from '10-Sep-2025' to '10/09/2025'"""
    months = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    
    parts = re.split(r'[-]', date_str.lower())
    if len(parts) == 3:
        day, month_abbr, year = parts
        month_num = months.get(month_abbr[:3], '01')
        return f"{day.zfill(2)}/{month_num}/{year}"
    return date_str

def extract_nirvana_fields(input_data, email_subject=""):
    """
    Extract reservation fields from Nirvana email content
    
    Args:
        input_data (str): Either email body text or path to .msg file
        email_subject (str): Email subject (optional)
    
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
        'MAIL_C_T_S': 'Nirvana',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    try:
        # Check if input_data is a file path or email text
        if input_data.endswith('.msg') and os.path.exists(input_data):
            # Handle .msg file
            msg = extract_msg.Message(input_data)
            email_body = msg.body or ""
            if not email_subject:
                email_subject = msg.subject or ""
        else:
            # Handle direct email text
            email_body = input_data
        
        if not email_body:
            return fields
        
        # Extract names - Nirvana specific patterns (Ms Nazira Nazir format)
        name_patterns = [
            r'(Ms|Mr|Mrs)\s+([A-Z][a-z]+\s+[A-Z][a-z]+)',
            r'1\s*Room\s*\n\s*([A-Z][a-z]+\s+[A-Z][a-z]+)',  # "1 Room \n Ms Nazira Nazir"
            r'Passengers[^:]*Room[^A-Z]*([A-Z][a-z]+\s+[A-Z][a-z]+)'
        ]
        
        for pattern in name_patterns:
            match = re.search(pattern, email_body, re.MULTILINE | re.DOTALL)
            if match:
                if match.lastindex >= 2 and match.group(2):  # Title + Name pattern
                    name_text = match.group(2).strip()
                else:
                    name_text = match.group(1).strip()
                
                # Split into first and full name
                name_parts = name_text.split()
                if len(name_parts) >= 2:  # Make sure we have at least first and last name
                    fields['MAIL_FIRST_NAME'] = name_parts[0]
                    fields['MAIL_FULL_NAME'] = name_text
                    break
        
        # Extract dates - Nirvana format (10-Sep-2025)
        arrival_patterns = [
            r'Check In\s*(\d{1,2}[-]\w{3}[-]\d{4})',
            r'Arrival Date[:\s]*(\d{2}/\d{2}/\d{4})',
            r'Check[- ]?in[:\s]*(\d{2}/\d{2}/\d{4})',
            r'From[:\s]*(\d{2}/\d{2}/\d{4})'
        ]
        
        for pattern in arrival_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # Convert 10-Sep-2025 to 10/09/2025
                if re.match(r'\d{1,2}[-]\w{3}[-]\d{4}', date_str):
                    date_str = convert_nirvana_date(date_str)
                fields['MAIL_ARRIVAL'] = date_str
                break
        
        departure_patterns = [
            r'Check Out\s*(\d{1,2}[-]\w{3}[-]\d{4})',
            r'Departure Date[:\s]*(\d{2}/\d{2}/\d{4})',
            r'Check[- ]?out[:\s]*(\d{2}/\d{2}/\d{4})',
            r'To[:\s]*(\d{2}/\d{2}/\d{4})'
        ]
        
        for pattern in departure_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # Convert 15-Sep-2025 to 15/09/2025
                if re.match(r'\d{1,2}[-]\w{3}[-]\d{4}', date_str):
                    date_str = convert_nirvana_date(date_str)
                fields['MAIL_DEPARTURE'] = date_str
                break
        
        # Extract nights directly or calculate from dates
        nights_match = re.search(r'No\.\s*of\s*Nights\s*(\d+)', email_body, re.IGNORECASE)
        if nights_match:
            fields['MAIL_NIGHTS'] = int(nights_match.group(1))
        elif fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
            try:
                arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
                dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
                nights = (dep_date - arr_date).days
                fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
            except:
                try:
                    # Try alternative format
                    arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%m/%d/%Y')
                    dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%m/%d/%Y')
                    nights = (dep_date - arr_date).days
                    fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
                except:
                    fields['MAIL_NIGHTS'] = 1
        
        # Extract persons/guests (1 Pax format)
        person_patterns = [
            r'(\d+)\s*Pax',
            r'(\d+)\s*Room',  # "1 Room" indicates 1 person typically
            r'(\d+)\s*Adult',
            r'(\d+)\s*Guest',
            r'(\d+)\s*Person',
            r'Guests?[:\s]*(\d+)',
            r'Adults?[:\s]*(\d+)'
        ]
        
        for pattern in person_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                fields['MAIL_PERSONS'] = int(match.group(1))
                break
        
        # Extract room type (SUPERIOR ROOM format)
        room_patterns = [
            r'Room Type\s*([A-Z\s]+ROOM[^-]*)',
            r'([A-Z\s]+ROOM)\s*-',
            r'Room Type[:\s]*([A-Za-z\s\(\)]+(?:Suite|Room|Apartment|Studio))',
            r'Accommodation[:\s]*([A-Za-z\s\(\)]+(?:Suite|Room|Apartment|Studio))'
        ]
        
        raw_room_type = 'N/A'
        for pattern in room_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                raw_room_type = match.group(1).strip()
                break
        
        # Apply room type mapping based on official room mapping
        if raw_room_type != 'N/A':
            room_type_lower = raw_room_type.lower()
            
            # Map according to official "Entered On room Map.xlsx"
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
            elif ('club room with one king bed' in room_type_lower or 
                  ('club' in room_type_lower and 'king' in room_type_lower)):
                fields['MAIL_ROOM'] = 'CK'  # Club King
            elif ('club room with two twin beds' in room_type_lower or 
                  ('club' in room_type_lower and 'twin' in room_type_lower)):
                fields['MAIL_ROOM'] = 'CT'  # Club Twin
            elif 'studio' in room_type_lower:
                fields['MAIL_ROOM'] = 'SA'  # Studio Apartment
            elif 'one bedroom apartment' in room_type_lower or '1 bedroom' in room_type_lower:
                fields['MAIL_ROOM'] = '1BA'  # One Bedroom Apartment
            elif 'two bedroom apartment' in room_type_lower or '2 bedroom' in room_type_lower:
                fields['MAIL_ROOM'] = '2BA'  # Two Bedroom Apartment
            elif 'business suite' in room_type_lower:
                fields['MAIL_ROOM'] = 'BS'  # Business Suite
            elif 'executive suite' in room_type_lower:
                fields['MAIL_ROOM'] = 'ES'  # Executive Suite
            elif 'presidential suite' in room_type_lower:
                fields['MAIL_ROOM'] = 'PRES'  # Presidential Suite
            elif 'royal suite' in room_type_lower:
                fields['MAIL_ROOM'] = 'RS'  # Royal Suite
            elif 'superior' in room_type_lower:
                # Default Superior Room - assume King if not specified
                fields['MAIL_ROOM'] = 'SK'  # Superior King (default)
            elif 'deluxe' in room_type_lower:
                # Default Deluxe Room - assume King if not specified
                fields['MAIL_ROOM'] = 'DK'  # Deluxe King (default)
            elif 'club' in room_type_lower:
                # Default Club Room - assume King if not specified
                fields['MAIL_ROOM'] = 'CK'  # Club King (default)
            else:
                # If no mapping found, use the raw room type
                fields['MAIL_ROOM'] = raw_room_type
        
        # Extract rate code/promo code (Offer Note:TOBBJN format)
        rate_patterns = [
            r'Offer Note[:\s]*([A-Z0-9\s\{\}]+)',
            r'Rate Code[:\s]*([A-Z0-9\s\{\}]+)',
            r'Promo[:\s]*([A-Z0-9\s\{\}]+)',
            r'Code[:\s]*([A-Z0-9\s\{\}]+)',
            r'Reference[:\s]*([A-Z0-9\s\{\}]+)'
        ]
        
        for pattern in rate_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                promo_text = match.group(1).strip()
                # Clean up - take only the code part before any parentheses
                promo_text = re.split(r'\s*\)', promo_text)[0]
                fields['MAIL_RATE_CODE'] = promo_text.strip()
                break
        
        # Extract monetary values - Nirvana format (Total Charges AED 1000.000)
        amount_patterns = [
            r'Total Charges\s*AED\s*([0-9,]+\.?\d*)',
            r'Total[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Amount[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Cost[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Price[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)'
        ]
        
        net_total = 0
        for pattern in amount_patterns:
            match = re.search(pattern, email_body, re.IGNORECASE)
            if match:
                try:
                    net_total = float(match.group(1).replace(',', ''))
                    fields['MAIL_NET_TOTAL'] = net_total
                    break
                except ValueError:
                    continue
        
        # Calculate TDF based on room type and nights
        tdf = 0
        nights = fields['MAIL_NIGHTS'] if fields['MAIL_NIGHTS'] != 'N/A' else 1
        
        if fields['MAIL_ROOM'] != 'N/A':
            room = fields['MAIL_ROOM']
            is_two_bedroom = '2BA' in room.upper() or 'Two Bedroom' in room or 'Suite' in room
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
        
    except Exception as e:
        print(f"Error processing Nirvana email: {e}")
        return fields
    
    return fields

def is_nirvana_email(sender_email, text):
    """
    Check if email is from Nirvana
    
    Args:
        sender_email (str): Sender email address
        text (str): Email content
    
    Returns:
        bool: True if this is a Nirvana email
    """
    content = (sender_email + " " + text).lower()
    return (
        'nirvana' in content or
        'booking confirmed' in content or
        'sb25' in content or
        'confirmation' in content
    )

# Test function
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        result = extract_nirvana_fields(file_path)
        
        print("Nirvana Email Extraction Test:")
        print("=" * 50)
        for key, value in result.items():
            if isinstance(value, float):
                print(f"{key}: AED {value:.2f}")
            else:
                print(f"{key}: {value}")
    else:
        print("Usage: python nirvana_parser.py <msg_file_path>")