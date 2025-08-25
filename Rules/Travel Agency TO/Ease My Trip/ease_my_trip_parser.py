"""
Ease My Trip Email Parser
Extracts reservation data from Ease My Trip confirmation emails
Based on Dubai Link parser logic but adapted for Ease My Trip format
"""

import re
from datetime import datetime

def extract_ease_my_trip_fields(email_body, email_subject):
    """
    Extract reservation fields from Ease My Trip email content
    
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
        'MAIL_C_T_S': 'Ease My Trip',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_AMOUNT': 'N/A',
        'MAIL_ADR': 'N/A'
    }
    
    # Extract names - Ease My Trip specific format
    first_name_match = re.search(r'Name:\s*([A-Za-z\s]+)', email_body)
    last_name_match = re.search(r'Last Name:\s*([A-Za-z\s]+)', email_body)
    
    # For Ease My Trip: MAIL_FIRST_NAME = Name field, MAIL_FULL_NAME = Last Name field
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
    
    # Extract number of nights directly from email
    nights_match = re.search(r'No Of Nights:\s*(\d+)', email_body)
    if nights_match:
        fields['MAIL_NIGHTS'] = int(nights_match.group(1))
    
    # Calculate nights if not found directly (fallback method)
    if fields['MAIL_NIGHTS'] == 'N/A' and fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
        try:
            arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], '%d/%m/%Y')
            dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], '%d/%m/%Y')
            nights = (dep_date - arr_date).days
            fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
        except:
            fields['MAIL_NIGHTS'] = 1
    
    # Extract persons (adults)
    persons_match = re.search(r'Number of adults:\s*(\d+)', email_body)
    if persons_match:
        fields['MAIL_PERSONS'] = int(persons_match.group(1))
    
    # Extract room type - Ease My Trip specific format
    room_match = re.search(r'Rooms:\s*(.*?)(?:\s*Meal plan|\s*$)', email_body, re.DOTALL)
    if room_match:
        room_info = room_match.group(1).strip()
        # Clean up the room info
        room_info = re.sub(r'\s+', ' ', room_info)
        fields['MAIL_ROOM'] = room_info
    
    # Alternative room extraction from room type line
    room_type_match = re.search(r'Superior Room.*?\(([^)]+)\)', email_body)
    if room_type_match and fields['MAIL_ROOM'] == 'N/A':
        meal_plan_match = re.search(r'Meal plan:\s*([^\\n]+)', email_body)
        meal_plan = meal_plan_match.group(1).strip() if meal_plan_match else "N/A"
        fields['MAIL_ROOM'] = f"Superior Room ({room_type_match.group(1)}) - {meal_plan}"
    
    # Extract promo code (rate code) - from LEISURE PROMOTION section
    leisure_promo_match = re.search(r'LEISURE PROMOTION.*?Promo code:\s*([A-Z0-9]+)', email_body)
    if leisure_promo_match:
        fields['MAIL_RATE_CODE'] = leisure_promo_match.group(1).strip()
    
    # Alternative promo code extraction 
    promo_match = re.search(r'Promo Code:\s*([^)]+)', email_body)
    if promo_match and fields['MAIL_RATE_CODE'] == 'N/A':
        promo_code = promo_match.group(1).strip()
        # Only use if not "Without applied promotions"
        if "Without applied promotions" not in promo_code:
            fields['MAIL_RATE_CODE'] = promo_code
    
    # Extract cost price (net total)
    cost_match = re.search(r'Cost price:\s*([\d,.]+)\s*AED', email_body)
    net_total = 0
    if cost_match:
        net_total = float(cost_match.group(1).replace(',', ''))
        fields['MAIL_NET_TOTAL'] = net_total
    
    # Calculate TDF based on room type and nights (same logic as Dubai Link)
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
    
    # Calculate derived values (same logic as Dubai Link)
    if net_total > 0:
        mail_total = net_total + tdf
        mail_amount = net_total / 1.225
        mail_adr = mail_amount / nights if nights > 0 else 0
        
        fields['MAIL_TOTAL'] = mail_total
        fields['MAIL_AMOUNT'] = mail_amount
        fields['MAIL_ADR'] = mail_adr
    
    return fields

def is_ease_my_trip_email(sender_email, subject):
    """
    Check if email is from Ease My Trip
    
    Args:
        sender_email (str): Sender email address
        subject (str): Email subject
    
    Returns:
        bool: True if this is an Ease My Trip email
    """
    return (
        'easemytrip.com' in sender_email.lower() or 
        'emtstays.com' in sender_email.lower() or
        'ease my trip' in subject.lower() or
        ('booking' in subject.lower() and 'g5fp7c' in subject.lower())
    )

# Test function
if __name__ == "__main__":
    # Test with sample data extracted from the actual email
    sample_email = """
    Booking Code: G5FP7C
    Name: Hamad
    Last Name: Almubarak
    Booking date: 25/08/2025
    
    Hotel: Grand Millennium Dubai
    Number of adults: 2
    Children: 0
    Babies: 0
    
    Arrival Date: 27/08/2025
    Departure Date: 29/08/2025
    No Of Nights: 2
    
    Rooms: 1 x Superior Room (2 Adults)
    Meal plan: Bed&Breakfast
    
    Superior Room (Bed&Breakfast)
    
    Promo Code: Without applied promotions Without applied promotions
    LEISURE PROMOTION (Promo code: TOBBJN) (0.00)
    
    Cost price: 400.00 AED
    """
    
    result = extract_ease_my_trip_fields(sample_email, "Paid Booking wth Ref. No. G5FP7C")
    
    print("Ease My Trip Extraction Test:")
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