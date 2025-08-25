"""
Agoda Email Parser for INNLINK2WAY System
Handles specific extraction and calculation logic for T-Agoda reservations
"""

import re
import pandas as pd
from datetime import datetime
from typing import Dict, Any, Optional

class AgodaParser:
    """Parser specifically for Agoda (T-Agoda) emails via INNLINK2WAY"""
    
    def __init__(self):
        """Initialize with Agoda specific regex patterns"""
        self.patterns = {
            'GUEST_NAME_FULL': re.compile(r"Guest Name:\s*(.+?)(?:\n|Address:|$)", re.IGNORECASE),
            'FIRST_NAME': re.compile(r"Guest Name:\s*([^\s]+)", re.IGNORECASE),
            'LAST_NAME': re.compile(r"Guest Name:\s*[^\s]+\s+(.+?)(?:\n|Address:|$)", re.IGNORECASE),
            'ARRIVAL': re.compile(r"Arrive:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
            'DEPARTURE': re.compile(r"Depart:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
            'NIGHTS': re.compile(r"Total Nights\s*(\d+)\s*night", re.IGNORECASE),
            'PERSONS': re.compile(r"Adult/Children:\s*(\d+)/\d+", re.IGNORECASE),
            'ROOM_TYPE': re.compile(r"Room Type:\s*(.+?)(?:\n|Rate|$)|(Superior Room|Deluxe Room|Standard Room|Executive Room|Studio with One King Bed)", re.IGNORECASE),
            'RATE_CODE': re.compile(r"Rate Code:\s*([A-Z0-9]+)", re.IGNORECASE),
            'RATE_NAME': re.compile(r"Rate Name:\s*(.+?)(?:\n|Rate Code:|$)", re.IGNORECASE),
            'COMPANY': re.compile(r"Travel Agent\s*(?:.*\n)*Name:\s*(.+?)(?:\n|$)", re.IGNORECASE | re.DOTALL),
            'NET_TOTAL_CHARGES': re.compile(r"Total charges:\s*AED\s*([0-9,]+\.?[0-9]*)", re.IGNORECASE),
            'CONFIRMATION': re.compile(r"Confirman:\s*([A-Z0-9]+)", re.IGNORECASE),
            'ARRIVAL_SUBJECT': re.compile(r"Arrival Date[:]*\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
            'CONFIRMATION_SUBJECT': re.compile(r"confirmation number[:]*\s*([A-Z0-9]+)", re.IGNORECASE),
        }
    
    def parse_agoda_email(self, email_content: str, sender_email: str = "") -> Dict[str, Any]:
        """
        Parse Agoda email content and extract reservation fields
        Apply T-Agoda specific business logic
        """
        extracted = {}
        
        # Extract all basic fields
        for field, pattern in self.patterns.items():
            match = pattern.search(email_content)
            if match:
                # Handle regex patterns with multiple groups
                for i in range(1, len(match.groups()) + 1):
                    if match.group(i) is not None:
                        extracted[field] = match.group(i).strip()
                        break
                else:
                    extracted[field] = "N/A"
            else:
                extracted[field] = "N/A"
        
        # Process guest names - Extract FIRST_NAME and FULL_NAME (last name)
        guest_name_full = extracted.get('GUEST_NAME_FULL', 'N/A')
        if guest_name_full != 'N/A' and guest_name_full.strip():
            name_parts = guest_name_full.strip().split()
            if len(name_parts) >= 2:
                extracted['FIRST_NAME'] = name_parts[0]
                extracted['FULL_NAME'] = name_parts[-1]  # Last name as FULL_NAME
            else:
                extracted['FIRST_NAME'] = guest_name_full
                extracted['FULL_NAME'] = guest_name_full
        else:
            extracted['FIRST_NAME'] = 'N/A'
            extracted['FULL_NAME'] = 'N/A'
        
        # Map room types to codes based on official room mapping
        room_type = extracted.get('ROOM_TYPE', 'N/A')
        if room_type != 'N/A':
            # Use official room mapping from "Entered On room Map.xlsx"
            if 'Superior Room with One King Bed' in room_type or ('Superior' in room_type and 'King' in room_type):
                extracted['ROOM'] = 'SK'  # Superior Room with One King Bed
            elif 'Superior Room with Two Twin Beds' in room_type or ('Superior' in room_type and 'Twin' in room_type):
                extracted['ROOM'] = 'ST'  # Superior Room with Two Twin Beds
            elif 'Deluxe Room with One King Bed' in room_type or ('Deluxe' in room_type and 'King' in room_type):
                extracted['ROOM'] = 'DK'  # Deluxe Room with One King Bed
            elif 'Deluxe Room with Two Twin Beds' in room_type or ('Deluxe' in room_type and 'Twin' in room_type):
                extracted['ROOM'] = 'DT'  # Deluxe Room with Two Twin Beds
            elif 'Club Room with One King Bed' in room_type or ('Club' in room_type and 'King' in room_type):
                extracted['ROOM'] = 'CK'  # Club Room with One King Bed
            elif 'Club Room with Two Twin Beds' in room_type or ('Club' in room_type and 'Twin' in room_type):
                extracted['ROOM'] = 'CT'  # Club Room with Two Twin Beds
            elif 'Studio with One King Bed' in room_type or 'Studio' in room_type:
                extracted['ROOM'] = 'SA'  # Studio with One King Bed
            elif 'One Bedroom Apartment' in room_type or '1 Bedroom' in room_type:
                extracted['ROOM'] = '1BA'  # One Bedroom Apartment
            elif 'Business Suite' in room_type:
                extracted['ROOM'] = 'BS'  # Business Suite with One King Bed
            elif 'Executive Suite' in room_type:
                extracted['ROOM'] = 'ES'  # Executive Suite with One King Bed
            elif 'Family Suite' in room_type:
                extracted['ROOM'] = 'FS'  # Family Suite
            elif 'Two Bedroom Apartment' in room_type or '2 Bedroom' in room_type:
                extracted['ROOM'] = '2BA'  # Two Bedroom Apartment
            elif 'Presidential Suite' in room_type:
                extracted['ROOM'] = 'PRES'  # Presidential Suite
            elif 'Royal Suite' in room_type:
                extracted['ROOM'] = 'RS'  # Royal Suite
            else:
                extracted['ROOM'] = room_type[:4].upper().replace(' ', '')
        else:
            extracted['ROOM'] = 'N/A'
        
        # Use rate code as primary, fallback to rate name
        if extracted.get('RATE_CODE', 'N/A') != 'N/A':
            pass  # Keep rate code
        elif extracted.get('RATE_NAME', 'N/A') != 'N/A':
            extracted['RATE_CODE'] = extracted['RATE_NAME']
        else:
            extracted['RATE_CODE'] = 'N/A'
        
        # Convert dates from mm/dd/yyyy to dd/mm/yyyy (INNLINK2WAY format)
        for date_field in ['ARRIVAL', 'DEPARTURE', 'ARRIVAL_SUBJECT']:
            if date_field in extracted and extracted[date_field] != 'N/A':
                try:
                    original_date = extracted[date_field]
                    # Parse as mm/dd/yyyy and convert to dd/mm/yyyy
                    parsed_date = pd.to_datetime(original_date, dayfirst=False)
                    extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
                except:
                    pass  # Keep original if parsing fails
        
        # Use arrival from subject if main arrival not found
        if extracted.get('ARRIVAL', 'N/A') == 'N/A' and extracted.get('ARRIVAL_SUBJECT', 'N/A') != 'N/A':
            extracted['ARRIVAL'] = extracted['ARRIVAL_SUBJECT']
        
        # Set company information - Agoda specific
        extracted['C_T_S'] = "T- Agoda"
        extracted['C_T_S_NAME'] = "T- Agoda"
        extracted['COMPANY'] = "T- Agoda"
        
        # ** T-AGODA SPECIFIC AMOUNT CALCULATIONS **
        net_total_charges = extracted.get('NET_TOTAL_CHARGES', 'N/A')
        nights = extracted.get('NIGHTS', 'N/A')
        
        if net_total_charges != 'N/A' and nights != 'N/A':
            try:
                # Parse net total charges (this is MAIL_NET_TOTAL for T-Agoda - amount WITHOUT TDF)
                net_total_amount = float(str(net_total_charges).replace(',', ''))
                nights_num = int(nights)
                
                # Calculate TDF (nights × 20)
                tdf_amount = nights_num * 20
                extracted['TDF'] = str(tdf_amount)
                
                # For T-Agoda: Email amount is MAIL_NET_TOTAL (excludes TDF)
                extracted['NET_TOTAL'] = str(net_total_amount)
                
                # MAIL_TOTAL = MAIL_NET_TOTAL + MAIL_TDF (total amount with TDF)
                total_with_tdf = net_total_amount + tdf_amount
                extracted['TOTAL'] = str(total_with_tdf)
                
                # MAIL_AMOUNT = MAIL_NET_TOTAL / 1.225 (amount without taxes)
                amount_without_taxes = net_total_amount / 1.225
                extracted['AMOUNT'] = f"{amount_without_taxes:.2f}"
                
                # Calculate ADR = AMOUNT / NIGHTS
                if nights_num > 0:
                    adr = amount_without_taxes / nights_num
                    extracted['ADR'] = f"{adr:.2f}"
                    extracted['ADR_AED'] = f"AED {adr:,.2f}"
                else:
                    extracted['ADR'] = "N/A"
                
                # Format currency fields
                extracted['TDF_AED'] = f"AED {tdf_amount:,.2f}"
                extracted['NET_TOTAL_AED'] = f"AED {net_total_amount:,.2f}"
                extracted['TOTAL_AED'] = f"AED {total_with_tdf:,.2f}"
                extracted['AMOUNT_AED'] = f"AED {amount_without_taxes:,.2f}"
                
            except (ValueError, TypeError):
                # If calculation fails, set defaults
                extracted['TDF'] = "N/A"
                extracted['NET_TOTAL'] = "N/A"
                extracted['TOTAL'] = "N/A"
                extracted['AMOUNT'] = "N/A"
                extracted['ADR'] = "N/A"
        else:
            extracted['TDF'] = "N/A"
            extracted['NET_TOTAL'] = "N/A"
            extracted['TOTAL'] = "N/A" 
            extracted['AMOUNT'] = "N/A"
            extracted['ADR'] = "N/A"
        
        return extracted
    
    def test_extraction(self, email_content: str, sender_email: str = "") -> Dict[str, Any]:
        """Test the extraction and return formatted results"""
        extracted = self.parse_agoda_email(email_content, sender_email)
        
        # Format for display/testing
        test_fields = [
            'FIRST_NAME', 'FULL_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 
            'ROOM', 'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT'
        ]
        
        results = {}
        for field in test_fields:
            value = extracted.get(field, 'N/A')
            # Format currency fields for display
            if field in ['NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT'] and value != 'N/A':
                try:
                    amount = float(str(value).replace(',', ''))
                    results[f'MAIL_{field}'] = f"AED {amount:,.2f}"
                except:
                    results[f'MAIL_{field}'] = value
            else:
                results[f'MAIL_{field}'] = value
        
        return results

def test_agoda_parser():
    """Test function for the Agoda parser"""
    print("Agoda Parser Test")
    print("=" * 50)
    
    # Sample email content (you can replace with actual content)
    sample_content = """
    Guest Name: JOHN SMITH
    Arrive: 06/09/2025
    Depart: 10/09/2025
    Total Nights 4 nights
    Adult/Children: 2/0
    Room Type: Superior Room with One King Bed
    Rate Code: AG123456
    Rate Name: Agoda Special Rate
    Travel Agent
    Name: Agoda
    Total charges: AED 2,400.00
    Confirman: 4K76RPPXK
    """
    
    parser = AgodaParser()
    results = parser.test_extraction(sample_content, "noreply-reservations@millenniumhotels.com")
    
    for field, value in results.items():
        print(f"{field:20}: {value}")
    
    return results

if __name__ == "__main__":
    test_agoda_parser()