"""
Miracle Tourism Parser
Extracts reservation data from Miracle Tourism booking forms (PDF files or .msg files with PDF attachments)
"""

import re
import pdfplumber
import extract_msg
import os
import tempfile
from datetime import datetime

def convert_month_format(date_str):
    """Convert date format from '01-Oct-25' to '01/10/2025'"""
    months = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    
    parts = re.split(r'[-\/]', date_str.lower())
    if len(parts) == 3:
        day, month_abbr, year = parts
        month_num = months.get(month_abbr[:3], '01')
        # Convert 2-digit year to 4-digit year
        if len(year) == 2:
            year = '20' + year
        return f"{day.zfill(2)}/{month_num}/{year}"
    return date_str

def extract_miracle_tourism_fields(file_path):
    """
    Extract reservation fields from Miracle Tourism booking form (.pdf or .msg file)
    
    Args:
        file_path (str): Path to the PDF file or .msg file
    
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
        'MAIL_C_T_S': 'Miracle Tourism',
        'MAIL_NET_TOTAL': 'N/A',
        'MAIL_TOTAL': 'N/A',
        'MAIL_TDF': 'N/A',
        'MAIL_ADR': 'N/A',
        'MAIL_AMOUNT': 'N/A'
    }
    
    try:
        pdf_text = ""
        
        # Handle different file types
        if file_path.lower().endswith('.msg'):
            # Extract from .msg file - check for PDF attachments or email body
            msg = extract_msg.Message(file_path)
            
            # First try to extract PDF attachments
            pdf_found = False
            with tempfile.TemporaryDirectory() as temp_dir:
                if msg.attachments:
                    for attachment in msg.attachments:
                        if attachment.longFilename and attachment.longFilename.lower().endswith('.pdf'):
                            pdf_path = os.path.join(temp_dir, attachment.longFilename)
                            with open(pdf_path, 'wb') as f:
                                f.write(attachment.data)
                            
                            # Extract text from PDF attachment
                            with pdfplumber.open(pdf_path) as pdf:
                                for page in pdf.pages:
                                    pdf_text += page.extract_text() or ""
                            pdf_found = True
                            break
                
                # If no PDF attachment found, try email body
                if not pdf_found:
                    if msg.body:
                        pdf_text = msg.body
                    elif msg.htmlBody:
                        # Simple HTML to text conversion
                        pdf_text = re.sub(r'<[^>]+>', '', msg.htmlBody)
        
        elif file_path.lower().endswith('.pdf'):
            # Direct PDF file
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    pdf_text += page.extract_text() or ""
        
        if not pdf_text:
            return fields
        
        # Extract first name and last name (Miracle Tourism format)
        name_patterns = [
            r'Names?[:\s]*([A-Z]+\s+[A-Z]+)',  # Match exactly two names in caps
            r'Guest\s*Names?[:\s]*([A-Z][A-Za-z\s]+)',
            r'First\s*Name[:\s]*([A-Z][A-Za-z\s]+)',
            r'Customer[:\s]*([A-Z][A-Za-z\s]+)',
            r'Booking\s*Ref[:\s]*\d+\s*([A-Z][A-Za-z\s]+)',
            r'GRAND\s*MILLENNIUM\s*DUBAI\s*([A-Z][A-Za-z\s]+)'
        ]
        
        for pattern in name_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                # Extract first name and last name from full name
                full_name = match.group(1).strip()
                # Take only the first line if there are multiple lines
                first_line = full_name.split('\n')[0].strip()
                name_parts = first_line.split()
                if name_parts:
                    fields['MAIL_FIRST_NAME'] = name_parts[0]  # First name
                    if len(name_parts) > 1:
                        fields['MAIL_FULL_NAME'] = ' '.join(name_parts[1:])  # Last name
                break
        
        # Extract arrival date (Miracle Tourism format)
        arrival_patterns = [
            r'Check\s*In[:\s]*(\d{1,2}[-\/]\w{3}[-\/]\d{2,4})',
            r'Arrival[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'From[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'Check-in[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\s*-\s*\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}'
        ]
        
        for pattern in arrival_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # Convert formats like "01-Oct-25" to "01/10/2025"
                if re.match(r'\d{1,2}[-\/]\w{3}[-\/]\d{2,4}', date_str):
                    date_str = convert_month_format(date_str)
                else:
                    # Normalize date format to DD/MM/YYYY
                    date_str = re.sub(r'[-]', '/', date_str)
                fields['MAIL_ARRIVAL'] = date_str
                break
        
        # Extract departure date (Miracle Tourism format)
        departure_patterns = [
            r'Check\s*Out[:\s]*(\d{1,2}[-\/]\w{3}[-\/]\d{2,4})',
            r'Departure[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'To[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'Until[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'Check-out[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
            r'\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}\s*-\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})'
        ]
        
        for pattern in departure_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # Convert formats like "02-Oct-25" to "02/10/2025"
                if re.match(r'\d{1,2}[-\/]\w{3}[-\/]\d{2,4}', date_str):
                    date_str = convert_month_format(date_str)
                else:
                    # Normalize date format to DD/MM/YYYY
                    date_str = re.sub(r'[-]', '/', date_str)
                fields['MAIL_DEPARTURE'] = date_str
                break
        
        # Calculate nights
        if fields['MAIL_ARRIVAL'] != 'N/A' and fields['MAIL_DEPARTURE'] != 'N/A':
            try:
                # Try different date formats
                for date_format in ['%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%m-%d-%Y']:
                    try:
                        arr_date = datetime.strptime(fields['MAIL_ARRIVAL'], date_format)
                        dep_date = datetime.strptime(fields['MAIL_DEPARTURE'], date_format)
                        nights = (dep_date - arr_date).days
                        fields['MAIL_NIGHTS'] = nights if nights > 0 else 1
                        break
                    except ValueError:
                        continue
            except:
                fields['MAIL_NIGHTS'] = 1
        
        # Extract number of persons (Miracle Tourism format)
        person_patterns = [
            r'No\.\s*of\s*Adult\'?s?[:\s]*(\d+)',
            r'(\d+)\s*Adult',
            r'(\d+)\s*Guest',
            r'(\d+)\s*Person',
            r'Pax[:\s]*(\d+)',
            r'Guests?[:\s]*(\d+)',
            r'(\d+)\s*Night'  # Sometimes nights are mentioned instead
        ]
        
        for pattern in person_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                fields['MAIL_PERSONS'] = int(match.group(1))
                break
        
        # Extract room type (Miracle Tourism format)
        room_patterns = [
            r'Room\s*Type[:\s]*([A-Z\s\/]+)',
            r'Room[:\s]*([A-Za-z\s\(\)\/]+(?:Suite|Room|Apartment|Studio))',
            r'Accommodation[:\s]*([A-Za-z\s\(\)\/]+(?:Suite|Room|Apartment|Studio))',
            r'Type[:\s]*([A-Za-z\s\(\)\/]+(?:Suite|Room|Apartment|Studio))',
            r'GRAND\s*MILLENNIUM\s*DUBAI[^a-zA-Z]*([A-Za-z\s\(\)\/]+(?:Suite|Room|Apartment|Studio))'
        ]
        
        for pattern in room_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                room_text = match.group(1).strip()
                # Clean up the room text - remove extra parts
                room_text = re.split(r'\s*(?:Conf|Nights|Check)', room_text)[0]
                fields['MAIL_ROOM'] = room_text.strip()
                break
        
        # Extract rate code or promo code (Miracle Tourism format)
        rate_patterns = [
            r'Promotions?[:\s]*([A-Z0-9\s\{\}]+)',
            r'Rate\s*Code[:\s]*([A-Z0-9\s\{\}]+)',
            r'Promo[:\s]*([A-Z0-9\s\{\}]+)',
            r'Code[:\s]*([A-Z0-9\s\{\}]+)',
            r'Booking\s*Ref[:\s]*([A-Z0-9]+)'
        ]
        
        for pattern in rate_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                promo_text = match.group(1).strip()
                # Clean up - take only the code part before any parentheses
                promo_text = re.split(r'\s*\(', promo_text)[0]
                fields['MAIL_RATE_CODE'] = promo_text.strip()
                break
        
        # Extract monetary values
        amount_patterns = [
            r'Total[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Amount[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Cost[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)',
            r'Price[:\s]*(?:AED\s*)?([0-9,]+\.?\d*)'
        ]
        
        for pattern in amount_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
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
        if fields['MAIL_NET_TOTAL'] != 'N/A' and fields['MAIL_NET_TOTAL'] > 0:
            net_total = fields['MAIL_NET_TOTAL']
            mail_total = net_total + tdf
            mail_amount = net_total / 1.225
            mail_adr = mail_amount / nights if nights > 0 else 0
            
            fields['MAIL_TOTAL'] = mail_total
            fields['MAIL_AMOUNT'] = mail_amount
            fields['MAIL_ADR'] = mail_adr
        
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return fields
    
    return fields

def is_miracle_tourism_file(file_path):
    """
    Check if file is from Miracle Tourism
    
    Args:
        file_path (str): Path to the file (.pdf or .msg)
    
    Returns:
        bool: True if this is a Miracle Tourism file
    """
    try:
        if file_path.lower().endswith('.msg'):
            msg = extract_msg.Message(file_path)
            content = ""
            if msg.body:
                content += msg.body.lower()
            if msg.subject:
                content += msg.subject.lower()
            return (
                'miracle tourism' in content or
                'miracle' in content or
                'luxair booking' in content or
                'booking ref' in content
            )
        elif file_path.lower().endswith('.pdf'):
            with pdfplumber.open(file_path) as pdf:
                first_page_text = pdf.pages[0].extract_text() or ""
                return (
                    'miracle tourism' in first_page_text.lower() or
                    'miracle' in first_page_text.lower() or
                    'luxair booking' in first_page_text.lower() or
                    'booking ref' in first_page_text.lower()
                )
    except:
        return False
    return False

# Test function
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        result = extract_miracle_tourism_fields(file_path)
        
        print("Miracle Tourism File Extraction Test:")
        print("=" * 50)
        for key, value in result.items():
            if isinstance(value, float):
                print(f"{key}: AED {value:.2f}")
            else:
                print(f"{key}: {value}")
    else:
        print("Usage: python miracle_tourism_parser.py <file_path>")