"""
Voyage Email Parser
Extracts reservation data from Voyage confirmation emails
"""

import re
import os
import sys
from datetime import datetime

# Cross-platform MSG file handling
try:
    import extract_msg
    MSG_LIBRARY_AVAILABLE = True
except ImportError:
    MSG_LIBRARY_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def check_room_count_and_extract_totals(email_body):
    """
    Check if booking has one or two rooms and extract individual totals
    
    Args:
        email_body (str): Email body content
    
    Returns:
        dict: Room information with count and individual totals
    """
    
    # Look for room count in the accommodation section only (not supplements)
    rooms_section_match = re.search(r'Rooms:\s*(.*?)(?=Allotment info:|$)', email_body, re.DOTALL)
    
    room_info = {
        'room_count': 0,
        'room_totals': [],
        'total_amount': 0,
        'room_descriptions': []
    }
    
    if rooms_section_match:
        rooms_text = rooms_section_match.group(1)
        # Extract actual room bookings from the Rooms section
        room_descriptions = re.findall(r'(\d+) x ([^-\n]+ - [^(\n]+(?:\([^)]+\))?)', rooms_text)
        
        if room_descriptions:
            room_info['room_count'] = len(room_descriptions)
            room_info['room_descriptions'] = [f"{count} x {desc.strip()}" for count, desc in room_descriptions]
    
    # Extract individual room totals from the detailed tables
    # Pattern: Look for amounts after room charges and child supplements
    table_sections = re.split(r'Superior Room - Double\s*\n', email_body)
    
    for i, section in enumerate(table_sections[1:], 1):  # Skip first section (header)
        # Look for room total + child supplement in each section
        room_charge_match = re.search(r'Room\s+[\d,.]+\s+x\s+\d+\s+([\d,.]+)', section)
        child_charge_match = re.search(r'2nd range child\s+[\d,.]+\s+x\s+\d+\s+([\d,.]+)', section)
        
        room_total = 0
        if room_charge_match:
            room_total += float(room_charge_match.group(1).replace(',', ''))
        if child_charge_match:
            room_total += float(child_charge_match.group(1).replace(',', ''))
        
        if room_total > 0 and len(room_info['room_totals']) < 2:  # Limit to 2 rooms max
            room_info['room_totals'].append(room_total)
    
    # If we didn't find individual totals, try to extract from simple pattern
    if not room_info['room_totals']:
        all_room_charges = re.findall(r'Room\s+[\d,.]+\s+x\s+\d+\s+([\d,.]+)', email_body)
        room_info['room_totals'] = [float(charge.replace(',', '')) for charge in all_room_charges[:2]]  # Limit to 2
    
    # Ensure room_count matches actual bookings
    if room_info['room_totals']:
        room_info['room_count'] = len(room_info['room_totals'])
    
    # Calculate total
    room_info['total_amount'] = sum(room_info['room_totals'])
    
    return room_info

def read_msg_file_cross_platform(msg_path):
    """
    Read .msg file content using cross-platform approach
    
    Args:
        msg_path (str): Path to .msg file
    
    Returns:
        dict: Email data with subject, sender, body etc.
    """
    
    if MSG_LIBRARY_AVAILABLE:
        # Use extract-msg library (cross-platform)
        try:
            msg = extract_msg.Message(msg_path)
            email_data = {
                'subject': msg.subject or '',
                'sender': msg.sender or '',
                'sender_name': msg.sender or '',
                'body': msg.body or '',
                'received_time': str(msg.date) if msg.date else '',
                'attachments': []
            }
            
            # Process attachments
            for attachment in msg.attachments:
                email_data['attachments'].append({
                    'filename': attachment.longFilename or attachment.shortFilename or '',
                    'type': 'pdf' if (attachment.longFilename or '').lower().endswith('.pdf') else 'other'
                })
            
            return email_data
            
        except Exception as e:
            print(f"Error reading .msg file with extract-msg: {e}")
            
    elif WIN32_AVAILABLE:
        # Fallback to Windows COM objects
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            msg = outlook.Session.OpenSharedItem(msg_path)
            
            email_data = {
                'subject': getattr(msg, 'Subject', ''),
                'sender': getattr(msg, 'SenderEmailAddress', ''),
                'sender_name': getattr(msg, 'SenderName', ''),
                'body': getattr(msg, 'Body', ''),
                'received_time': str(getattr(msg, 'ReceivedTime', '')),
                'attachments': []
            }
            
            if hasattr(msg, 'Attachments') and msg.Attachments.Count > 0:
                for attachment in msg.Attachments:
                    filename = getattr(attachment, 'FileName', '')
                    email_data['attachments'].append({
                        'filename': filename,
                        'type': 'pdf' if filename.lower().endswith('.pdf') else 'other'
                    })
            
            return email_data
            
        except Exception as e:
            print(f"Error reading .msg file with win32com: {e}")
            return None
        finally:
            pythoncom.CoUninitialize()
    
    else:
        print("No MSG reading libraries available. Install extract-msg or run on Windows with pywin32.")
        return None

def extract_voyage_fields_from_msg(msg_path):
    """
    Extract Voyage fields directly from .msg file - returns separate room data
    
    Args:
        msg_path (str): Path to .msg file
    
    Returns:
        list: List of extracted field values for each room
    """
    
    # Read the MSG file
    email_data = read_msg_file_cross_platform(msg_path)
    
    if not email_data:
        return None
    
    # Extract fields from the email content
    email_body = email_data['body']
    email_subject = email_data['subject']
    
    return extract_voyage_fields_separate_rooms(email_body, email_subject)

def extract_voyage_fields_separate_rooms(email_body, email_subject):
    """
    Extract reservation fields from Voyage email content - separate for each room
    
    Args:
        email_body (str): Email body content
        email_subject (str): Email subject
    
    Returns:
        list: List of extracted field dictionaries, one per room
    """
    
    # Get room information
    room_info = check_room_count_and_extract_totals(email_body)
    
    # Extract common information (same for all rooms)
    common_fields = extract_common_voyage_fields(email_body, email_subject)
    
    # Create separate entries for each room
    room_entries = []
    
    for i, room_total in enumerate(room_info['room_totals']):
        room_fields = common_fields.copy()
        
        # Update room-specific fields
        if i < len(room_info['room_descriptions']):
            room_fields['MAIL_ROOM'] = room_info['room_descriptions'][i]
        else:
            room_fields['MAIL_ROOM'] = f"Room {i+1}"
        
        # Individual room costs
        room_fields['MAIL_NET_TOTAL'] = room_total
        
        # Individual TDF (20 AED per room per night)
        nights = room_fields['MAIL_NIGHTS'] if room_fields['MAIL_NIGHTS'] != 'N/A' else 1
        effective_nights = min(nights, 30) if nights >= 30 else nights
        room_tdf = effective_nights * 20  # 20 AED per room per night
        room_fields['MAIL_TDF'] = room_tdf
        
        # Individual totals
        mail_total = room_total + room_tdf
        mail_amount = room_total / 1.225 if room_total > 0 else 0
        mail_adr = mail_amount / nights if nights > 0 and mail_amount > 0 else 0
        
        room_fields['MAIL_TOTAL'] = mail_total
        room_fields['MAIL_AMOUNT'] = mail_amount
        room_fields['MAIL_ADR'] = mail_adr
        
        # Add room number identifier
        room_fields['ROOM_NUMBER'] = i + 1
        
        room_entries.append(room_fields)
    
    return room_entries

def extract_common_voyage_fields(email_body, email_subject):
    """
    Extract common fields that are the same for all rooms
    
    Args:
        email_body (str): Email body content
        email_subject (str): Email subject
    
    Returns:
        dict: Common extracted field values
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
        'MAIL_C_T_S': 'Voyage',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract names - Voyage specific mapping (try multiple patterns)
    first_name_match = re.search(r'Name:\s*([A-Z\s]+)', email_body)
    last_name_match = re.search(r'Last Name:\s*([A-Z\s]+)', email_body)
    contact_person_match = re.search(r'Contact person\s+([^%\n]+)', email_body)
    nationality_match = re.search(r'Nationality:\s*([A-Z\s]+)', email_body)
    
    # Look for passenger information - first passenger
    passenger_match = re.search(r'1\.\s*([A-Z]+)\s+([A-Z]+)\s*\([^)]+\)', email_body)
    
    # For Voyage: Try different sources for name
    if passenger_match:
        # Extract from passenger list: "1. ADEL ALAZMI (30 age)"
        fields['MAIL_FIRST_NAME'] = passenger_match.group(1).strip()  # ADEL
        fields['MAIL_FULL_NAME'] = passenger_match.group(2).strip()   # ALAZMI
    elif first_name_match:
        fields['MAIL_FIRST_NAME'] = first_name_match.group(1).strip()
        if last_name_match:
            fields['MAIL_FULL_NAME'] = last_name_match.group(1).strip()
    elif contact_person_match:
        fields['MAIL_FIRST_NAME'] = contact_person_match.group(1).strip()
        if nationality_match:
            fields['MAIL_FULL_NAME'] = nationality_match.group(1).strip()
    
    # Extract dates - Updated patterns for Voyage format
    arrival_match = re.search(r'Check-In Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    departure_match = re.search(r'Check-Out Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    
    # Fallback patterns
    if not arrival_match:
        arrival_match = re.search(r'Arrival Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    if not departure_match:
        departure_match = re.search(r'Departure Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    
    if arrival_match:
        fields['MAIL_ARRIVAL'] = arrival_match.group(1)
    if departure_match:
        fields['MAIL_DEPARTURE'] = departure_match.group(1)
    
    # Extract nights directly from email or calculate
    nights_match = re.search(r'Nights:\s*(\d+)', email_body)
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
    
    # Extract persons - count adults from room descriptions
    adults_matches = re.findall(r'(\d+) Adults?', email_body)
    if adults_matches:
        # Sum all adults from all rooms
        total_adults = sum(int(match) for match in adults_matches)
        fields['MAIL_PERSONS'] = total_adults
    else:
        # Fallback pattern
        persons_match = re.search(r'\((\d+) Adult\)', email_body)
        if persons_match:
            fields['MAIL_PERSONS'] = int(persons_match.group(1))
    
    # Extract promo code (rate code)
    promo_match = re.search(r'Promo code:\s*([A-Z0-9{}\s]+)', email_body)
    if promo_match:
        fields['MAIL_RATE_CODE'] = promo_match.group(1).strip()
    
    return fields

def extract_voyage_fields(email_body, email_subject):
    """
    Extract reservation fields from Voyage email content
    
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
        'MAIL_C_T_S': 'Voyage',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract names - Voyage specific mapping (try multiple patterns)
    first_name_match = re.search(r'Name:\s*([A-Z\s]+)', email_body)
    last_name_match = re.search(r'Last Name:\s*([A-Z\s]+)', email_body)
    contact_person_match = re.search(r'Contact person\s+([^%\n]+)', email_body)
    nationality_match = re.search(r'Nationality:\s*([A-Z\s]+)', email_body)
    
    # Look for passenger information - first passenger
    passenger_match = re.search(r'1\.\s*([A-Z]+)\s+([A-Z]+)\s*\([^)]+\)', email_body)
    
    # For Voyage: Try different sources for name
    if passenger_match:
        # Extract from passenger list: "1. ADEL ALAZMI (30 age)"
        fields['MAIL_FIRST_NAME'] = passenger_match.group(1).strip()  # ADEL
        fields['MAIL_FULL_NAME'] = passenger_match.group(2).strip()   # ALAZMI
    elif first_name_match:
        fields['MAIL_FIRST_NAME'] = first_name_match.group(1).strip()
        if last_name_match:
            fields['MAIL_FULL_NAME'] = last_name_match.group(1).strip()
    elif contact_person_match:
        fields['MAIL_FIRST_NAME'] = contact_person_match.group(1).strip()
        if nationality_match:
            fields['MAIL_FULL_NAME'] = nationality_match.group(1).strip()
    
    # Extract dates - Updated patterns for Voyage format
    arrival_match = re.search(r'Check-In Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    departure_match = re.search(r'Check-Out Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    
    # Fallback patterns
    if not arrival_match:
        arrival_match = re.search(r'Arrival Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    if not departure_match:
        departure_match = re.search(r'Departure Date:\s*(\d{2}/\d{2}/\d{4})', email_body)
    
    if arrival_match:
        fields['MAIL_ARRIVAL'] = arrival_match.group(1)
    if departure_match:
        fields['MAIL_DEPARTURE'] = departure_match.group(1)
    
    # Extract nights directly from email or calculate
    nights_match = re.search(r'Nights:\s*(\d+)', email_body)
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
    
    # Extract persons - count adults from room descriptions
    adults_matches = re.findall(r'(\d+) Adults?', email_body)
    if adults_matches:
        # Sum all adults from all rooms
        total_adults = sum(int(match) for match in adults_matches)
        fields['MAIL_PERSONS'] = total_adults
    else:
        # Fallback pattern
        persons_match = re.search(r'\((\d+) Adult\)', email_body)
        if persons_match:
            fields['MAIL_PERSONS'] = int(persons_match.group(1))
    
    # Extract room type - capture all room descriptions
    room_matches = re.findall(r'(\d+ x [^-\n]+ - [^(\n]+(?:\([^)]+\))?)', email_body)
    if room_matches:
        # Filter out duplicate or partial matches and clean them up
        unique_rooms = []
        for room in room_matches:
            cleaned = room.strip()
            # Skip if this looks like a partial match or duplicate
            if len(cleaned) > 20 and cleaned not in unique_rooms:
                unique_rooms.append(cleaned)
        
        if unique_rooms:
            fields['MAIL_ROOM'] = '\n'.join(unique_rooms[:2])  # Limit to first 2 unique rooms
    
    if fields['MAIL_ROOM'] == 'N/A':
        # Fallback pattern
        room_match = re.search(r'(\d+ x [^(]+\([^)]+\)[^)]*)', email_body)
        if room_match:
            fields['MAIL_ROOM'] = room_match.group(1).strip()
    
    # Extract promo code (rate code)
    promo_match = re.search(r'Promo code:\s*([A-Z0-9{}\s]+)', email_body)
    if promo_match:
        fields['MAIL_RATE_CODE'] = promo_match.group(1).strip()
    
    # Check room count and extract individual totals
    room_info = check_room_count_and_extract_totals(email_body)
    
    # Extract booking cost using room information
    net_total = 0
    
    if room_info['total_amount'] > 0:
        net_total = room_info['total_amount']
        fields['MAIL_NET_TOTAL'] = net_total
        
        # Update room description with cleaner format
        if room_info['room_descriptions']:
            fields['MAIL_ROOM'] = '\n'.join(room_info['room_descriptions'])
            
        # Add room breakdown info for debugging
        if len(room_info['room_totals']) > 1:
            print(f"Room breakdown: Room 1: AED {room_info['room_totals'][0]:.2f}, Room 2: AED {room_info['room_totals'][1]:.2f}")
    
    else:
        # Fallback extraction methods
        # Try pattern: amount before "Total"
        total_before_match = re.search(r'([\d,.]+)\s*\n?\s*Total', email_body)
        if total_before_match:
            net_total = float(total_before_match.group(1).replace(',', ''))
            fields['MAIL_NET_TOTAL'] = net_total
        else:
            # Try summing room charges from table
            room_charges = re.findall(r'Room\s+[\d,.]+\s+x\s+\d+\s+([\d,.]+)', email_body)
            if room_charges:
                net_total = sum(float(charge.replace(',', '')) for charge in room_charges)
                fields['MAIL_NET_TOTAL'] = net_total
            else:
                # Final fallback pattern
                cost_match = re.search(r'Booking cost price:\s*([\d,.]+)\s*AED', email_body)
                if cost_match:
                    net_total = float(cost_match.group(1).replace(',', ''))
                    fields['MAIL_NET_TOTAL'] = net_total
    
    # Calculate TDF based on room count and nights  
    tdf = 0
    nights = fields['MAIL_NIGHTS'] if fields['MAIL_NIGHTS'] != 'N/A' else 1
    
    # Use room count from room_info if available
    if 'room_info' in locals() and room_info['room_count'] > 0:
        room_count = room_info['room_count']
    elif fields['MAIL_ROOM'] != 'N/A':
        # Fallback: count room lines
        room = fields['MAIL_ROOM']
        room_count = len(room.split('\n')) if '\n' in room else 1
    else:
        room_count = 1
    
    # TDF is 20 AED per room per night for standard rooms
    tdf_rate_per_room = 20
    
    # For 30+ nights, use 30 as the multiplier instead of actual nights
    effective_nights = min(nights, 30) if nights >= 30 else nights
    
    # Calculate TDF for all rooms
    tdf = room_count * effective_nights * tdf_rate_per_room
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

def is_voyage_email(sender_email, subject):
    """
    Check if email is from Voyage
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is a Voyage email
    """
    return (
        'voyage' in sender_email.lower() or 
        'voyage' in subject.lower() or
        'booking request' in subject.lower()
    )

# Test function
if __name__ == "__main__":
    # Test with sample email content
    sample_email = """
    Name: CUSTOMER
    Last Name: NAME
    
    Arrival Date: 27/08/2025
    Departure Date: 28/08/2025
    Rooms: 1 x Superior Room (King/Twin) - Double (1 Adult)
    
    Booking cost price: 200.00 AED
    
    Promo code: TOBBJN{ALL MARKET EX UAE}
    """
    
    result = extract_voyage_fields(sample_email, "Booking Request with Ref. No. CZN277")
    
    print("Voyage Extraction Test:")
    print("=" * 50)
    for key, value in result.items():
        if isinstance(value, float):
            print(f"{key}: AED {value:.2f}")
        else:
            print(f"{key}: {value}")
    
    # Test with actual MSG file if provided as command line argument
    if len(sys.argv) > 1:
        msg_file_path = sys.argv[1]
        if os.path.exists(msg_file_path):
            print(f"\n\nTesting MSG file extraction: {msg_file_path}")
            print("=" * 80)
            
            msg_results = extract_voyage_fields_from_msg(msg_file_path)
            if msg_results and isinstance(msg_results, list):
                for i, room_result in enumerate(msg_results):
                    print(f"\n=== ROOM {i+1} EXTRACTION RESULTS ===")
                    print("-" * 50)
                    
                    # Display in vertical format
                    field_order = [
                        'MAIL_FIRST_NAME', 'MAIL_FULL_NAME', 'MAIL_ARRIVAL', 
                        'MAIL_DEPARTURE', 'MAIL_NIGHTS', 'MAIL_PERSONS',
                        'MAIL_ROOM', 'MAIL_RATE_CODE', 'MAIL_C_T_S',
                        'MAIL_NET_TOTAL', 'MAIL_TDF', 'MAIL_TOTAL',
                        'MAIL_AMOUNT', 'MAIL_ADR'
                    ]
                    
                    for key in field_order:
                        if key in room_result:
                            value = room_result[key]
                            if isinstance(value, float):
                                print(f"{key:20}: AED {value:.2f}")
                            else:
                                print(f"{key:20}: {value}")
                    print("-" * 50)
                    
                print(f"\nTotal rooms processed: {len(msg_results)}")
                
            elif msg_results:
                # Fallback for single room (old format)
                print("MSG Extraction Results:")
                print("-" * 50)
                for key, value in msg_results.items():
                    if isinstance(value, float):
                        print(f"{key}: AED {value:.2f}")
                    else:
                        print(f"{key}: {value}")
            else:
                print("Failed to extract data from MSG file")
        else:
            print(f"MSG file not found: {msg_file_path}")