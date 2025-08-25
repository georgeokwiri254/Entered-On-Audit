"""
Streamlit App for Entered On Audit System
Three tabs: Email Extraction Results, Converted Data, and Audit
Uses win32com.client to access Outlook emails locally
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import logging
import sqlite3
import json
import re
import pdfplumber
import io
from pathlib import Path
import win32com.client
import pythoncom

# Import our existing converter and database operations
from entered_on_converter import process_entered_on_report, get_summary_stats
from database_operations import AuditDatabase

# ** RULE ENGINE FOR TRAVEL AGENCY DETECTION **
def get_travel_agency_rule(c_t_s_name, sender_email="", text=""):
    """
    Determine which parser rule to use based on Travel Agency C_T_S name and content
    Returns: (rule_type, parser_path, insert_user)
    """
    
    # Clean the C_T_S name for comparison
    c_t_s_clean = str(c_t_s_name).strip() if c_t_s_name else ""
    
    # INNLINKWAY Rules - for C_T_S names starting with "T-"
    if c_t_s_clean.startswith("T-") or "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        insert_user = "*INNLINK2WAY*"
        
        # T-Agoda
        if ("agoda" in c_t_s_clean.lower() or 
            "agoda" in text.lower() or 
            "t- agoda" in text.lower()):
            return ("INNLINKWAY_AGODA", "Rules/INNLINKWAY/Agoda", insert_user)
        
        # T-Booking.com
        elif ("booking.com" in c_t_s_clean.lower() or 
              "booking.com" in text.lower() or 
              "t- booking.com" in text.lower()):
            return ("INNLINKWAY_BOOKING", "Rules/INNLINKWAY/Booking.com", insert_user)
        
        # T-Brand.com
        elif ("brand.com" in c_t_s_clean.lower() or 
              "brand.com" in text.lower() or 
              "t- brand.com" in text.lower()):
            return ("INNLINKWAY_BRAND", "Rules/INNLINKWAY/Brand.com", insert_user)
        
        # T-Expedia
        elif ("expedia" in c_t_s_clean.lower() or 
              "expedia" in text.lower() or 
              "t- expedia" in text.lower()):
            return ("INNLINKWAY_EXPEDIA", "Rules/INNLINKWAY/Expedia", insert_user)
        
        # Default INNLINKWAY rule (fallback to Brand.com logic)
        else:
            return ("INNLINKWAY_DEFAULT", "Rules/INNLINKWAY/Brand.com", insert_user)
    
    # Travel Agency Rules - Traditional travel agencies
    elif c_t_s_clean:
        insert_user = c_t_s_clean  # Use actual company name as INSERT_USER
        
        # Travco
        if ("travco" in c_t_s_clean.lower() or 
            "travco.co.uk" in sender_email.lower() or
            "hotel booking confirmation" in text.lower()):
            return ("TRAVEL_AGENCY_TRAVCO", "Rules/Travel Agency TO/Travco", insert_user)
        
        # Dubai Link
        elif ("dubai link" in c_t_s_clean.lower() or 
              "gte.travel" in sender_email.lower() or
              "dubai link" in text.lower()):
            return ("TRAVEL_AGENCY_DUBAI_LINK", "Rules/Travel Agency TO/Dubai Link", insert_user)
        
        # Nirvana
        elif ("nirvana" in c_t_s_clean.lower() or 
              "nirvana" in sender_email.lower() or
              "booking confirmed" in text.lower()):
            return ("TRAVEL_AGENCY_NIRVANA", "Rules/Travel Agency TO/Nirvana", insert_user)
        
        # Dakkak DMC / Duri Travel
        elif ("dakkak" in c_t_s_clean.lower() or 
              "dakkak" in sender_email.lower() or
              "dakkak dmc" in text.lower()):
            return ("TRAVEL_AGENCY_DAKKAK", "Rules/Travel Agency TO/Dakkak", insert_user)
        
        # Duri
        elif ("duri" in c_t_s_clean.lower() or 
              "hanmail.net" in sender_email.lower() or
              "duri travel" in text.lower()):
            return ("TRAVEL_AGENCY_DURI", "Rules/Travel Agency TO/Duri", insert_user)
        
        # AlKhalidiah
        elif ("alkhalidiah" in c_t_s_clean.lower() or 
              "alkhalidiah.com" in sender_email.lower() or
              "al khalidiah" in text.lower()):
            return ("TRAVEL_AGENCY_ALKHALIDIAH", "Rules/Travel Agency TO/AlKhalidiah", insert_user)
        
        # Desert Adventures
        elif ("desert adventures" in c_t_s_clean.lower() or
              "allocation notification" in text.lower()):
            return ("TRAVEL_AGENCY_DESERT_ADVENTURES", "Rules/Travel Agency TO/Desert Adventures", insert_user)
        
        # Desert Gate
        elif ("desert gate" in c_t_s_clean.lower() or
              "dgt" in sender_email.lower() or
              "booking notification" in text.lower()):
            return ("TRAVEL_AGENCY_DESERT_GATE", "Rules/Travel Agency TO/Desert Gate", insert_user)
        
        # Darina
        elif ("darina" in c_t_s_clean.lower() or
              "booking form" in text.lower()):
            return ("TRAVEL_AGENCY_DARINA", "Rules/Travel Agency TO/Darina", insert_user)
        
        # Ease My Trip
        elif ("ease my trip" in c_t_s_clean.lower() or
              "paid booking" in text.lower()):
            return ("TRAVEL_AGENCY_EASE_MY_TRIP", "Rules/Travel Agency TO/Ease My Trip", insert_user)
        
        # Almosafer
        elif ("almosafer" in c_t_s_clean.lower() or
              "confirmed booking" in text.lower()):
            return ("TRAVEL_AGENCY_ALMOSAFER", "Rules/Travel Agency TO/Almosafer", insert_user)
        
        # Generic Travel Agency - fallback
        else:
            return ("TRAVEL_AGENCY_GENERIC", None, insert_user)
    
    # Airlines Rules
    elif ("china southern" in text.lower() or 
          "c- china southern" in text.lower()):
        return ("AIRLINES_CHINA_SOUTHERN", "Rules/Airlines/China Air", "China Southern Air")
    
    # UPS Airlines
    elif ("ups" in c_t_s_clean.lower() or 
          "ups" in text.lower()):
        return ("AIRLINES_UPS", "Rules/Airlines/UPS", "UPS Airlines")
    
    # ASL Airlines
    elif ("asl" in c_t_s_clean.lower() or 
          "asl" in text.lower()):
        return ("AIRLINES_ASL", "Rules/Airlines/ASL", "ASL Airlines")
    
    # Corporate or Group Rate
    elif ("corporate" in c_t_s_clean.lower() or 
          "grp" in c_t_s_clean.lower()):
        return ("CORPORATE_RATE", "Rules/Corporate COR", c_t_s_clean)
    
    # Default - no specific rule
    else:
        return ("DEFAULT", None, "MANUAL_ENTRY")

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="Entered On Audit System",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'email_data' not in st.session_state:
    st.session_state.email_data = None
if 'audit_results' not in st.session_state:
    st.session_state.audit_results = None
if 'uploaded_file_name' not in st.session_state:
    st.session_state.uploaded_file_name = None
if 'selected_file_path' not in st.session_state:
    st.session_state.selected_file_path = None
if 'auto_loaded' not in st.session_state:
    st.session_state.auto_loaded = False
if 'current_run_id' not in st.session_state:
    st.session_state.current_run_id = None
if 'database' not in st.session_state:
    st.session_state.database = AuditDatabase()

# Helper functions for email processing using Outlook COM
def connect_to_outlook():
    """Connect to Outlook using win32com.client"""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        logger.error(f"Outlook connection failed: {e}")
        return None, None

def extract_pdf_text(pdf_bytes):
    """Extract text from PDF bytes with enhanced error handling and performance optimizations"""
    try:
        # Skip very large PDFs (> 5MB) to avoid performance issues
        if len(pdf_bytes) > 5 * 1024 * 1024:
            logger.warning(f"Skipping large PDF ({len(pdf_bytes) / (1024*1024):.1f}MB) - too large for processing")
            return ""
        
        pdf_file = io.BytesIO(pdf_bytes)
        text = ""
        
        with pdfplumber.open(pdf_file) as pdf:
            page_count = len(pdf.pages)
            logger.info(f"PDF has {page_count} pages")
            
            # Limit processing to first 3 pages for performance
            max_pages = min(3, page_count)
            
            for page_num, page in enumerate(pdf.pages[:max_pages]):
                try:
                    page_text = page.extract_text()
                    if page_text:
                        text += f"\n--- Page {page_num + 1} ---\n{page_text}"
                        
                        # Early exit if we found China Southern Air on first page
                        if page_num == 0 and ("china southern" in page_text.lower() or "c- china southern" in page_text.lower()):
                            logger.info("Found China Southern Air on first page - processing first page only")
                            break
                    else:
                        logger.warning(f"No text extracted from page {page_num + 1}")
                except Exception as e:
                    logger.warning(f"Failed to extract text from page {page_num + 1}: {e}")
                    continue
        
        if not text.strip():
            logger.warning("No text extracted from PDF - may be image-based")
            
        return text.strip()
        
    except Exception as e:
        logger.error(f"PDF extraction failed: {e}")
        return ""

# Compile regex patterns once for better performance
import re

# Pre-compiled patterns for different email types
NOREPLY_PATTERNS = {
    'GUEST_NAME_FULL': re.compile(r"Guest Name:\s*(.+?)(?:\n|Address:)", re.IGNORECASE),
    'ARRIVAL': re.compile(r"Arrive:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'DEPARTURE': re.compile(r"Depart:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'NIGHTS': re.compile(r"Total Nights\s*(\d+)\s*night", re.IGNORECASE),
    'PERSONS': re.compile(r"Adult/Children:\s*(\d+)/\d+", re.IGNORECASE),
    'ROOM_TYPE': re.compile(r"Room Type:\s*(.+?)(?:\n|Rate|$)|(Superior Room|Deluxe Room|Standard Room|Executive Room|Studio with One King Bed)", re.IGNORECASE),
    'RATE_CODE': re.compile(r"Rate Code:\s*([A-Z0-9]+)", re.IGNORECASE),
    'RATE_NAME': re.compile(r"Rate Name:\s*(.+?)(?:\n|Rate Code:)", re.IGNORECASE),
    'COMPANY': re.compile(r"Travel Agent\s*(?:.*\n)*Name:\s*(.+?)(?:\n|$)", re.IGNORECASE | re.DOTALL),
    'NET_TOTAL': re.compile(r"Total charges:\s*AED\s*([0-9,]+\.?[0-9]*)", re.IGNORECASE),
    'CONFIRMATION': re.compile(r"Confirman:\s*([A-Z0-9]+)", re.IGNORECASE),
    'ARRIVAL_SUBJECT': re.compile(r"Arrival Date[:]*\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'CONFIRMATION_SUBJECT': re.compile(r"confirmation number[:]*\s*([A-Z0-9]+)", re.IGNORECASE),
}

CHINA_SOUTHERN_PATTERNS = {
    'FULL_NAME': re.compile(r"(?:Passenger Name|Guest Name|Name)[:\s]*([A-Z][A-Za-z\s]+)(?:\n|Cabin|Flight)", re.IGNORECASE),
    'FIRST_NAME': re.compile(r"(?:First Name|Given Name)[:\s]*([A-Za-z]+)", re.IGNORECASE),
    'ARRIVAL': re.compile(r"(?:Arrival Date|Check.?in|Arrival)[:\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'DEPARTURE': re.compile(r"(?:Departure Date|Check.?out|Departure)[:\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'NIGHTS': re.compile(r"(?:Nights?|Night Stay|Duration)[:\s]*(\d+)", re.IGNORECASE),
    'PERSONS': re.compile(r"(?:Passengers?|Guests?|Adults?|Pax)[:\s]*(\d+)", re.IGNORECASE),
    'ROOM': re.compile(r"(?:Room Type|Cabin|Accommodation)[:\s]*([A-Z0-9\s]+)", re.IGNORECASE),
    'RATE_CODE': re.compile(r"(?:Rate Code|Booking Code|Reference)[:\s]*([A-Z0-9]+)", re.IGNORECASE),
    'NET_TOTAL': re.compile(r"(?:Total Cost|Total Amount|Net Total|Total)[:\s]*(?:AED|USD)?\s*([0-9,]+\.?[0-9]*)", re.IGNORECASE),
    'CONFIRMATION': re.compile(r"(?:PNR|Confirmation|Booking Reference)[:\s]*([A-Z0-9]+)", re.IGNORECASE),
    'FLIGHT': re.compile(r"(?:Flight|Flight Number)[:\s]*([A-Z0-9]+)", re.IGNORECASE),
}

DEFAULT_PATTERNS = {
    'FULL_NAME': re.compile(r"(?:Name|Guest Name)[:\s]+(.+?)(?:\n|$)", re.IGNORECASE),
    'FIRST_NAME': re.compile(r"(?:First Name)[:\s]+(.+?)(?:\n|$)", re.IGNORECASE),
    'ARRIVAL': re.compile(r"(?:Arrival|Check-in)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'DEPARTURE': re.compile(r"(?:Departure|Check-out)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", re.IGNORECASE),
    'NIGHTS': re.compile(r"(?:Nights|Night)[:\s]+(\d+)", re.IGNORECASE),
    'PERSONS': re.compile(r"(?:Persons|Guest|Adults?)[:\s]+(\d+)", re.IGNORECASE),
    'ROOM': re.compile(r"(?:Room|Room Type)[:\s]+(.+?)(?:\n|$)", re.IGNORECASE),
    'RATE_CODE': re.compile(r"(?:Rate Code|Rate)[:\s]+(.+?)(?:\n|$)", re.IGNORECASE),
    'COMPANY': re.compile(r"(?:Company|Agency)[:\s]+(.+?)(?:\n|$)", re.IGNORECASE),
    'NET_TOTAL': re.compile(r"(?:Total|Net Total|Amount|Net Amount)[:\s]+(?:AED\s*)?([\\d,]+\.?\\d*)", re.IGNORECASE),
}

def extract_reservation_fields(text, sender_email="", c_t_s_name=""):
    """Extract reservation fields using rule-based parser selection for better performance"""
    
    # Use rule engine to determine which parser to use
    rule_type, parser_path, insert_user = get_travel_agency_rule(c_t_s_name, sender_email, text)
    
    # Log the rule selection for debugging
    logger.info(f"Rule engine selected: {rule_type} for C_T_S: {c_t_s_name}, Email: {sender_email}")
    
    # Check for Travco emails first
    if "travco.co.uk" in sender_email.lower() or "travco@travco" in sender_email.lower() or "hotel booking confirmation" in text.lower():
        # Import Travco parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Travco'))
        try:
            from travco_parser import extract_travco_fields, is_travco_email
            
            if is_travco_email(sender_email, text):
                travco_fields = extract_travco_fields(text, "")
                # Map Travco fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': travco_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': travco_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': travco_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': travco_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': travco_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': travco_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': travco_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': travco_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': travco_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': travco_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': travco_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': travco_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': travco_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': travco_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': travco_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {travco_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if travco_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {travco_fields.get('MAIL_TOTAL', 0):,.2f}" if travco_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {travco_fields.get('MAIL_TDF', 0):,.2f}" if travco_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {travco_fields.get('MAIL_ADR', 0):,.2f}" if travco_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {travco_fields.get('MAIL_AMOUNT', 0):,.2f}" if travco_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("Travco parser not found, falling back to default patterns")
    
    # Check for Dubai Link emails
    if "gte.travel" in sender_email.lower() or "dubai link" in text.lower():
        # Import Dubai Link parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Dubai Link'))
        try:
            from dubai_link_parser import extract_dubai_link_fields, is_dubai_link_email
            
            if is_dubai_link_email(sender_email, text):
                dubai_fields = extract_dubai_link_fields(text, "")
                # Map Dubai Link fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': dubai_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': dubai_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': dubai_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': dubai_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': dubai_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': dubai_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': dubai_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': dubai_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': dubai_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': dubai_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': dubai_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': dubai_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': dubai_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': dubai_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': dubai_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {dubai_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if dubai_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {dubai_fields.get('MAIL_TOTAL', 0):,.2f}" if dubai_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {dubai_fields.get('MAIL_TDF', 0):,.2f}" if dubai_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {dubai_fields.get('MAIL_ADR', 0):,.2f}" if dubai_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {dubai_fields.get('MAIL_AMOUNT', 0):,.2f}" if dubai_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("Dubai Link parser not found, falling back to default patterns")
    
    # Check for Nirvana emails
    if "nirvana" in sender_email.lower() or "booking confirmed" in text.lower() or "sb25" in text.lower():
        # Import Nirvana parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Nirvana'))
        try:
            from nirvana_parser import extract_nirvana_fields, is_nirvana_email
            
            if is_nirvana_email(sender_email, text):
                nirvana_fields = extract_nirvana_fields(text, "")
                # Map Nirvana fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': nirvana_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': nirvana_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': nirvana_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': nirvana_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': nirvana_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': nirvana_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': nirvana_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': nirvana_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': nirvana_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': nirvana_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': nirvana_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': nirvana_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': nirvana_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': nirvana_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': nirvana_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {nirvana_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if nirvana_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {nirvana_fields.get('MAIL_TOTAL', 0):,.2f}" if nirvana_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {nirvana_fields.get('MAIL_TDF', 0):,.2f}" if nirvana_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {nirvana_fields.get('MAIL_ADR', 0):,.2f}" if nirvana_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {nirvana_fields.get('MAIL_AMOUNT', 0):,.2f}" if nirvana_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("Nirvana parser not found, falling back to default patterns")
    
    # Check for Duri Travel / Dakkak DMC emails
    if "dakkak" in sender_email.lower() or "dakkak dmc" in text.lower() or "hotel new booking" in text.lower() and "bkgho" in text.lower():
        # Import Duri Travel parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Duri Travel'))
        try:
            from duri_travel_parser import extract_duri_travel_fields, is_duri_travel_email
            
            if is_duri_travel_email(sender_email, text):
                duri_fields = extract_duri_travel_fields(text, "")
                # Map Duri Travel fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': duri_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': duri_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': duri_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': duri_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': duri_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': duri_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': duri_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': duri_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': duri_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': duri_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': duri_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': duri_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': duri_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': duri_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': duri_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {duri_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if duri_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {duri_fields.get('MAIL_TOTAL', 0):,.2f}" if duri_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {duri_fields.get('MAIL_TDF', 0):,.2f}" if duri_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {duri_fields.get('MAIL_ADR', 0):,.2f}" if duri_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {duri_fields.get('MAIL_AMOUNT', 0):,.2f}" if duri_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("Duri Travel parser not found, falling back to default patterns")
    
    # Check for Duri emails
    if "hanmail.net" in sender_email.lower() or "duri travel" in text.lower() or ("grand millennium dubai" in text.lower() and "jmc57" in sender_email.lower()):
        # Import Duri parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'Duri'))
        try:
            from duri_parser import extract_duri_fields, is_duri_email
            
            if is_duri_email(sender_email, text):
                duri_fields = extract_duri_fields(text, "")
                # Map Duri fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': duri_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': duri_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': duri_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': duri_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': duri_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': duri_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': duri_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': duri_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': duri_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': duri_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': duri_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': duri_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': duri_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': duri_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': duri_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {duri_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if duri_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {duri_fields.get('MAIL_TOTAL', 0):,.2f}" if duri_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {duri_fields.get('MAIL_TDF', 0):,.2f}" if duri_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {duri_fields.get('MAIL_ADR', 0):,.2f}" if duri_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {duri_fields.get('MAIL_AMOUNT', 0):,.2f}" if duri_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("Duri parser not found, falling back to default patterns")
    
    # Check for AlKhalidiah Tourism emails
    if "alkhalidiah.com" in sender_email.lower() or "alkhalidiah" in text.lower() or "al khalidiah" in text.lower():
        # Import AlKhalidiah parser
        import sys
        import os
        sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'Travel Agency TO', 'AlKhalidiah'))
        try:
            from alkhalidiah_parser import extract_alkhalidiah_fields, is_alkhalidiah_email
            
            if is_alkhalidiah_email(sender_email, text):
                alkhalidiah_fields = extract_alkhalidiah_fields(text, "")
                # Map AlKhalidiah fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': alkhalidiah_fields.get('MAIL_FIRST_NAME', 'N/A'),
                    'FULL_NAME': alkhalidiah_fields.get('MAIL_FULL_NAME', 'N/A'),
                    'ARRIVAL': alkhalidiah_fields.get('MAIL_ARRIVAL', 'N/A'),
                    'DEPARTURE': alkhalidiah_fields.get('MAIL_DEPARTURE', 'N/A'),
                    'NIGHTS': alkhalidiah_fields.get('MAIL_NIGHTS', 'N/A'),
                    'PERSONS': alkhalidiah_fields.get('MAIL_PERSONS', 'N/A'),
                    'ROOM': alkhalidiah_fields.get('MAIL_ROOM', 'N/A'),
                    'RATE_CODE': alkhalidiah_fields.get('MAIL_RATE_CODE', 'N/A'),
                    'C_T_S': alkhalidiah_fields.get('MAIL_C_T_S', 'N/A'),
                    'C_T_S_NAME': alkhalidiah_fields.get('MAIL_C_T_S', 'N/A'),
                    'NET_TOTAL': alkhalidiah_fields.get('MAIL_NET_TOTAL', 'N/A'),
                    'TOTAL': alkhalidiah_fields.get('MAIL_TOTAL', 'N/A'),
                    'TDF': alkhalidiah_fields.get('MAIL_TDF', 'N/A'),
                    'ADR': alkhalidiah_fields.get('MAIL_ADR', 'N/A'),
                    'AMOUNT': alkhalidiah_fields.get('MAIL_AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': f"AED {alkhalidiah_fields.get('MAIL_NET_TOTAL', 0):,.2f}" if alkhalidiah_fields.get('MAIL_NET_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TOTAL_AED': f"AED {alkhalidiah_fields.get('MAIL_TOTAL', 0):,.2f}" if alkhalidiah_fields.get('MAIL_TOTAL', 'N/A') != 'N/A' else 'N/A',
                    'TDF_AED': f"AED {alkhalidiah_fields.get('MAIL_TDF', 0):,.2f}" if alkhalidiah_fields.get('MAIL_TDF', 'N/A') != 'N/A' else 'N/A',
                    'ADR_AED': f"AED {alkhalidiah_fields.get('MAIL_ADR', 0):,.2f}" if alkhalidiah_fields.get('MAIL_ADR', 'N/A') != 'N/A' else 'N/A',
                    'AMOUNT_AED': f"AED {alkhalidiah_fields.get('MAIL_AMOUNT', 0):,.2f}" if alkhalidiah_fields.get('MAIL_AMOUNT', 'N/A') != 'N/A' else 'N/A'
                }
                return mapped_fields
        except ImportError:
            logger.warning("AlKhalidiah parser not found, falling back to default patterns")
    
    # ** INNLINKWAY PARSERS INTEGRATION **
    # Check for INNLINKWAY emails (noreply-reservations@millenniumhotels.com)
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        
        # T-Agoda parser
        if ("agoda" in text.lower() or "t- agoda" in text.lower() or "confirmation number" in text.lower()):
            import sys
            import os
            sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'INNLINKWAY', 'Agoda'))
            try:
                from agoda_parser import AgodaParser
                
                parser = AgodaParser()
                agoda_fields = parser.parse_agoda_email(text, sender_email)
                # Map Agoda fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': agoda_fields.get('FIRST_NAME', 'N/A'),
                    'FULL_NAME': agoda_fields.get('FULL_NAME', 'N/A'),
                    'ARRIVAL': agoda_fields.get('ARRIVAL', 'N/A'),
                    'DEPARTURE': agoda_fields.get('DEPARTURE', 'N/A'),
                    'NIGHTS': agoda_fields.get('NIGHTS', 'N/A'),
                    'PERSONS': agoda_fields.get('PERSONS', 'N/A'),
                    'ROOM': agoda_fields.get('ROOM', 'N/A'),
                    'RATE_CODE': agoda_fields.get('RATE_CODE', 'N/A'),
                    'C_T_S': agoda_fields.get('C_T_S', 'N/A'),
                    'C_T_S_NAME': agoda_fields.get('C_T_S_NAME', 'N/A'),
                    'NET_TOTAL': agoda_fields.get('NET_TOTAL', 'N/A'),
                    'TOTAL': agoda_fields.get('TOTAL', 'N/A'),
                    'TDF': agoda_fields.get('TDF', 'N/A'),
                    'ADR': agoda_fields.get('ADR', 'N/A'),
                    'AMOUNT': agoda_fields.get('AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': agoda_fields.get('NET_TOTAL_AED', 'N/A'),
                    'TOTAL_AED': agoda_fields.get('TOTAL_AED', 'N/A'),
                    'TDF_AED': agoda_fields.get('TDF_AED', 'N/A'),
                    'ADR_AED': agoda_fields.get('ADR_AED', 'N/A'),
                    'AMOUNT_AED': agoda_fields.get('AMOUNT_AED', 'N/A')
                }
                return mapped_fields
            except ImportError:
                logger.warning("Agoda INNLINKWAY parser not found, falling back to default patterns")
        
        # T-Booking.com parser
        elif ("booking.com" in text.lower() or "t- booking.com" in text.lower()):
            import sys
            import os
            sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'INNLINKWAY', 'Booking.com'))
            try:
                from booking_com_parser import BookingComParser
                
                parser = BookingComParser()
                booking_fields = parser.parse_booking_email(text, sender_email)
                # Map Booking.com fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': booking_fields.get('FIRST_NAME', 'N/A'),
                    'FULL_NAME': booking_fields.get('FULL_NAME', 'N/A'),
                    'ARRIVAL': booking_fields.get('ARRIVAL', 'N/A'),
                    'DEPARTURE': booking_fields.get('DEPARTURE', 'N/A'),
                    'NIGHTS': booking_fields.get('NIGHTS', 'N/A'),
                    'PERSONS': booking_fields.get('PERSONS', 'N/A'),
                    'ROOM': booking_fields.get('ROOM', 'N/A'),
                    'RATE_CODE': booking_fields.get('RATE_CODE', 'N/A'),
                    'C_T_S': booking_fields.get('C_T_S', 'N/A'),
                    'C_T_S_NAME': booking_fields.get('C_T_S_NAME', 'N/A'),
                    'NET_TOTAL': booking_fields.get('NET_TOTAL', 'N/A'),
                    'TOTAL': booking_fields.get('TOTAL', 'N/A'),
                    'TDF': booking_fields.get('TDF', 'N/A'),
                    'ADR': booking_fields.get('ADR', 'N/A'),
                    'AMOUNT': booking_fields.get('AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': booking_fields.get('NET_TOTAL_AED', 'N/A'),
                    'TOTAL_AED': booking_fields.get('TOTAL_AED', 'N/A'),
                    'TDF_AED': booking_fields.get('TDF_AED', 'N/A'),
                    'ADR_AED': booking_fields.get('ADR_AED', 'N/A'),
                    'AMOUNT_AED': booking_fields.get('AMOUNT_AED', 'N/A')
                }
                return mapped_fields
            except ImportError:
                logger.warning("Booking.com INNLINKWAY parser not found, falling back to default patterns")
        
        # T-Brand.com parser
        elif ("brand.com" in text.lower() or "t- brand.com" in text.lower()):
            import sys
            import os
            sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'INNLINKWAY', 'Brand.com'))
            try:
                from brand_com_parser import BrandComParser
                
                parser = BrandComParser()
                brand_fields = parser.parse_brand_email(text, sender_email)
                # Map Brand.com fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': brand_fields.get('FIRST_NAME', 'N/A'),
                    'FULL_NAME': brand_fields.get('FULL_NAME', 'N/A'),
                    'ARRIVAL': brand_fields.get('ARRIVAL', 'N/A'),
                    'DEPARTURE': brand_fields.get('DEPARTURE', 'N/A'),
                    'NIGHTS': brand_fields.get('NIGHTS', 'N/A'),
                    'PERSONS': brand_fields.get('PERSONS', 'N/A'),
                    'ROOM': brand_fields.get('ROOM', 'N/A'),
                    'RATE_CODE': brand_fields.get('RATE_CODE', 'N/A'),
                    'C_T_S': brand_fields.get('C_T_S', 'N/A'),
                    'C_T_S_NAME': brand_fields.get('C_T_S_NAME', 'N/A'),
                    'NET_TOTAL': brand_fields.get('NET_TOTAL', 'N/A'),
                    'TOTAL': brand_fields.get('TOTAL', 'N/A'),
                    'TDF': brand_fields.get('TDF', 'N/A'),
                    'ADR': brand_fields.get('ADR', 'N/A'),
                    'AMOUNT': brand_fields.get('AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': brand_fields.get('NET_TOTAL_AED', 'N/A'),
                    'TOTAL_AED': brand_fields.get('TOTAL_AED', 'N/A'),
                    'TDF_AED': brand_fields.get('TDF_AED', 'N/A'),
                    'ADR_AED': brand_fields.get('ADR_AED', 'N/A'),
                    'AMOUNT_AED': brand_fields.get('AMOUNT_AED', 'N/A')
                }
                return mapped_fields
            except ImportError:
                logger.warning("Brand.com INNLINKWAY parser not found, falling back to default patterns")
        
        # T-Expedia parser
        elif ("expedia" in text.lower() or "t- expedia" in text.lower()):
            import sys
            import os
            sys.path.append(os.path.join(os.path.dirname(__file__), 'Rules', 'INNLINKWAY', 'Expedia'))
            try:
                from expedia_parser import ExpediaParser
                
                parser = ExpediaParser()
                expedia_fields = parser.parse_expedia_email(text, sender_email)
                # Map Expedia fields to the expected field names used in the app
                mapped_fields = {
                    'FIRST_NAME': expedia_fields.get('FIRST_NAME', 'N/A'),
                    'FULL_NAME': expedia_fields.get('FULL_NAME', 'N/A'),
                    'ARRIVAL': expedia_fields.get('ARRIVAL', 'N/A'),
                    'DEPARTURE': expedia_fields.get('DEPARTURE', 'N/A'),
                    'NIGHTS': expedia_fields.get('NIGHTS', 'N/A'),
                    'PERSONS': expedia_fields.get('PERSONS', 'N/A'),
                    'ROOM': expedia_fields.get('ROOM', 'N/A'),
                    'RATE_CODE': expedia_fields.get('RATE_CODE', 'N/A'),
                    'C_T_S': expedia_fields.get('C_T_S', 'N/A'),
                    'C_T_S_NAME': expedia_fields.get('C_T_S_NAME', 'N/A'),
                    'NET_TOTAL': expedia_fields.get('NET_TOTAL', 'N/A'),
                    'TOTAL': expedia_fields.get('TOTAL', 'N/A'),
                    'TDF': expedia_fields.get('TDF', 'N/A'),
                    'ADR': expedia_fields.get('ADR', 'N/A'),
                    'AMOUNT': expedia_fields.get('AMOUNT', 'N/A'),
                    # Add formatted currency versions
                    'NET_TOTAL_AED': expedia_fields.get('NET_TOTAL_AED', 'N/A'),
                    'TOTAL_AED': expedia_fields.get('TOTAL_AED', 'N/A'),
                    'TDF_AED': expedia_fields.get('TDF_AED', 'N/A'),
                    'ADR_AED': expedia_fields.get('ADR_AED', 'N/A'),
                    'AMOUNT_AED': expedia_fields.get('AMOUNT_AED', 'N/A')
                }
                return mapped_fields
            except ImportError:
                logger.warning("Expedia INNLINKWAY parser not found, falling back to default patterns")
    
    # Select pattern set based on email source for faster processing
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        patterns = NOREPLY_PATTERNS
    elif "c- china southern air" in text.lower() or "china southern" in text.lower():
        patterns = CHINA_SOUTHERN_PATTERNS  
    else:
        patterns = DEFAULT_PATTERNS
    
    extracted = {}
    
    # Extract all fields using pre-compiled patterns
    for field, compiled_pattern in patterns.items():
        match = compiled_pattern.search(text)
        if match:
            value = match.group(1).strip()
            extracted[field] = value
        else:
            extracted[field] = "N/A"
    
    # Special processing for noreply-reservations emails
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        # Process guest name - split "Boaz Avital" into first name and last name
        guest_name = extracted.get('GUEST_NAME_FULL', 'N/A')
        if guest_name != 'N/A' and guest_name.strip():
            name_parts = guest_name.strip().split()
            if len(name_parts) >= 2:
                # First name is the first part, full name (last name) is the last part
                extracted['FIRST_NAME'] = name_parts[0]
                extracted['FULL_NAME'] = name_parts[-1]  # Last name as full name per instruction
            else:
                extracted['FIRST_NAME'] = guest_name
                extracted['FULL_NAME'] = guest_name
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
            elif 'Business Suite with One King Bed' in room_type or ('Business Suite' in room_type):
                extracted['ROOM'] = 'BS'  # Business Suite with One King Bed
            elif 'Executive Suite with One King Bed' in room_type or ('Executive Suite' in room_type):
                extracted['ROOM'] = 'ES'  # Executive Suite with One King Bed
            elif 'Family Suite' in room_type:
                extracted['ROOM'] = 'FS'  # Family Suite with 1 King and 2 Twin Beds
            elif 'Two Bedroom Apartment' in room_type or '2 Bedroom' in room_type:
                extracted['ROOM'] = '2BA'  # Two Bedroom Apartment
            elif 'Presidential Suite' in room_type:
                extracted['ROOM'] = 'PRES'  # Presidential Suite
            elif 'Royal Suite' in room_type:
                extracted['ROOM'] = 'RS'  # Royal Suite
            else:
                # Fallback: try to extract first few characters
                extracted['ROOM'] = room_type[:4].upper().replace(' ', '')
        else:
            extracted['ROOM'] = 'N/A'
        
        # Use rate code as primary, fallback to rate name
        if extracted.get('RATE_CODE', 'N/A') != 'N/A':
            # Keep the rate code as is
            pass
        elif extracted.get('RATE_NAME', 'N/A') != 'N/A':
            extracted['RATE_CODE'] = extracted['RATE_NAME']
        else:
            extracted['RATE_CODE'] = 'N/A'
    
    # Convert dates to dd/mm/yyyy format - Enhanced logic for INNLINK2WAY
    for date_field in ['ARRIVAL', 'DEPARTURE', 'ARRIVAL_SUBJECT']:
        if date_field in extracted and extracted[date_field] != 'N/A':
            try:
                original_date = extracted[date_field]
                
                # Special handling for INNLINK2WAY and noreply-reservations emails
                if ("noreply-reservations@millenniumhotels.com" in sender_email.lower() or 
                    "innlink2way" in sender_email.lower()):
                    # For INNLINK2WAY, dates are typically in mm/dd/yyyy format that need conversion
                    try:
                        # First try parsing as mm/dd/yyyy (dayfirst=False)
                        parsed_date = pd.to_datetime(original_date, dayfirst=False)
                        extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
                        continue
                    except:
                        pass
                
                # Default: try dd/mm/yyyy format first
                try:
                    parsed_date = pd.to_datetime(original_date, dayfirst=True)
                    extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
                except:
                    # Fallback: try mm/dd/yyyy format
                    try:
                        parsed_date = pd.to_datetime(original_date, dayfirst=False)
                        extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
                    except:
                        pass  # Keep original value if all parsing fails
            except:
                pass
    
    # Use arrival from subject if main arrival not found
    if extracted.get('ARRIVAL', 'N/A') == 'N/A' and extracted.get('ARRIVAL_SUBJECT', 'N/A') != 'N/A':
        extracted['ARRIVAL'] = extracted['ARRIVAL_SUBJECT']
    
    # ** OTA-SPECIFIC CALCULATIONS **
    # Check if this is from INNLINKWAY (noreply-reservations@millenniumhotels.com)
    is_innlinkway = "noreply-reservations@millenniumhotels.com" in sender_email.lower()
    
    # Check if this is a T-Agoda or T-Expedia reservation (NET_TOTAL logic)
    is_agoda_expedia = ("agoda" in extracted.get('COMPANY', '').lower() or 
                       "agoda" in text.lower() or
                       "t- agoda" in text.lower() or
                       "expedia" in extracted.get('COMPANY', '').lower() or 
                       "expedia" in text.lower() or
                       "t- expedia" in text.lower())
    
    # Check if this should follow Booking.com logic (TOTAL logic)
    # Rule: Any INNLINKWAY reservation NOT from Agoda/Expedia follows Booking.com logic
    is_booking_logic = (is_innlinkway and not is_agoda_expedia) or (
                       "booking.com" in extracted.get('COMPANY', '').lower() or 
                       "booking.com" in text.lower() or
                       "t- booking.com" in text.lower())
    
    # Calculate TDF as nights × 20
    try:
        nights = extracted.get('NIGHTS', 'N/A')
        if nights != 'N/A' and str(nights).isdigit():
            nights_num = int(nights)
            tdf_amount = nights_num * 20
            extracted['TDF'] = str(tdf_amount)
            extracted['TDF_AED'] = f"AED {tdf_amount:,.2f}"
        else:
            extracted['TDF'] = "N/A"
            tdf_amount = 0
    except:
        extracted['TDF'] = "N/A"
        tdf_amount = 0
    
    # Handle amounts based on OTA type
    if is_booking_logic:
        # Booking.com Logic: Email amount is MAIL_TOTAL (includes TDF)
        # Applies to: T-Booking.com, Brand.com, and any INNLINKWAY non-Agoda/Expedia
        try:
            total_charges = extracted.get('NET_TOTAL', 'N/A')  # This is actually TOTAL for booking logic
            if total_charges != 'N/A':
                total_amount = float(str(total_charges).replace(',', ''))
                
                # For Booking Logic: MAIL_TOTAL = email amount (includes TDF)
                extracted['TOTAL'] = str(total_amount)
                
                # MAIL_NET_TOTAL = MAIL_TOTAL - MAIL_TDF (amount without TDF but with taxes)
                net_total = total_amount - tdf_amount
                extracted['NET_TOTAL'] = str(net_total)
                
                # MAIL_AMOUNT = MAIL_NET_TOTAL / 1.225 (amount without taxes)
                amount_without_taxes = net_total / 1.225
                extracted['AMOUNT'] = f"{amount_without_taxes:.2f}"
                
                # Calculate ADR = AMOUNT / NIGHTS
                if nights != 'N/A' and int(nights) > 0:
                    adr = amount_without_taxes / int(nights)
                    extracted['ADR'] = f"{adr:.2f}"
                    extracted['ADR_AED'] = f"AED {adr:,.2f}"
                else:
                    extracted['ADR'] = "N/A"
                
                # Format currency fields
                extracted['TOTAL_AED'] = f"AED {total_amount:,.2f}"
                extracted['NET_TOTAL_AED'] = f"AED {net_total:,.2f}"
                extracted['AMOUNT_AED'] = f"AED {amount_without_taxes:,.2f}"
            else:
                extracted['TOTAL'] = "N/A"
                extracted['AMOUNT'] = "N/A"
                extracted['ADR'] = "N/A"
        except:
            extracted['TOTAL'] = "N/A"
            extracted['AMOUNT'] = "N/A"
            extracted['ADR'] = "N/A"
    elif is_agoda_expedia and is_innlinkway:
        # T-Agoda/T-Expedia: Email amount is MAIL_NET_TOTAL (excludes TDF)
        try:
            net_charges = extracted.get('NET_TOTAL', 'N/A')  # This is NET_TOTAL for agoda/expedia
            if net_charges != 'N/A':
                net_total_amount = float(str(net_charges).replace(',', ''))
                
                # For T-Agoda/T-Expedia: MAIL_NET_TOTAL = email amount (excludes TDF)
                extracted['NET_TOTAL'] = str(net_total_amount)
                
                # MAIL_TOTAL = MAIL_NET_TOTAL + MAIL_TDF (total amount with TDF)
                total_with_tdf = net_total_amount + tdf_amount
                extracted['TOTAL'] = str(total_with_tdf)
                
                # MAIL_AMOUNT = MAIL_NET_TOTAL / 1.225 (amount without taxes)
                amount_without_taxes = net_total_amount / 1.225
                extracted['AMOUNT'] = f"{amount_without_taxes:.2f}"
                
                # Calculate ADR = AMOUNT / NIGHTS
                if nights != 'N/A' and int(nights) > 0:
                    adr = amount_without_taxes / int(nights)
                    extracted['ADR'] = f"{adr:.2f}"
                    extracted['ADR_AED'] = f"AED {adr:,.2f}"
                else:
                    extracted['ADR'] = "N/A"
                
                # Format currency fields
                extracted['NET_TOTAL_AED'] = f"AED {net_total_amount:,.2f}"
                extracted['TOTAL_AED'] = f"AED {total_with_tdf:,.2f}"
                extracted['AMOUNT_AED'] = f"AED {amount_without_taxes:,.2f}"
            else:
                extracted['NET_TOTAL'] = "N/A"
                extracted['TOTAL'] = "N/A"
                extracted['AMOUNT'] = "N/A"
                extracted['ADR'] = "N/A"
        except:
            extracted['NET_TOTAL'] = "N/A"
            extracted['TOTAL'] = "N/A"
            extracted['AMOUNT'] = "N/A"
            extracted['ADR'] = "N/A"
    else:
        # Default calculation for other OTAs (Expedia, Agoda, etc.)
        # Calculate ADR (Average Daily Rate) = NET_TOTAL / NIGHTS
        try:
            net_total = extracted.get('NET_TOTAL', 'N/A')
            nights = extracted.get('NIGHTS', 'N/A')
            if (net_total != 'N/A' and nights != 'N/A' and 
                str(nights).isdigit() and str(net_total).replace(',', '').replace('.', '').isdigit()):
                nights_num = int(nights)
                net_total_num = float(str(net_total).replace(',', ''))
                if nights_num > 0:
                    adr = net_total_num / nights_num
                    extracted['ADR'] = f"{adr:.2f}"
                    extracted['ADR_AED'] = f"AED {adr:,.2f}"
                else:
                    extracted['ADR'] = "N/A"
            else:
                extracted['ADR'] = "N/A"
        except:
            extracted['ADR'] = "N/A"
        
        # Set AMOUNT = NET_TOTAL for consistency (other OTAs)
        try:
            net_total = extracted.get('NET_TOTAL', 'N/A')
            if net_total != 'N/A':
                amount_num = float(str(net_total).replace(',', ''))
                extracted['AMOUNT'] = net_total
                extracted['AMOUNT_AED'] = f"AED {amount_num:,.2f}"
                # For non-booking.com, TOTAL = NET_TOTAL + TDF
                total_with_tdf = amount_num + tdf_amount
                extracted['TOTAL'] = str(total_with_tdf)
            else:
                extracted['AMOUNT'] = "N/A"
                extracted['TOTAL'] = "N/A"
        except:
            extracted['AMOUNT'] = "N/A"
            extracted['TOTAL'] = "N/A"
    
    # Special handling for China Southern Air reservations
    if "c- china southern air" in text.lower() or "china southern" in text.lower():
        extracted['C_T_S'] = "C- China Southern Air"
        extracted['C_T_S_NAME'] = "C- China Southern Air"
        extracted['COMPANY'] = "C- China Southern Air"
        
        # If we found flight info, store it in rate code
        if extracted.get('FLIGHT', 'N/A') != 'N/A':
            extracted['RATE_CODE'] = extracted['FLIGHT']
    else:
        # Map COMPANY to C_T_S (Company name) for other types
        if extracted.get('COMPANY', 'N/A') != 'N/A':
            extracted['C_T_S'] = extracted['COMPANY']
            extracted['C_T_S_NAME'] = extracted['COMPANY']
        else:
            extracted['C_T_S'] = "N/A"
            extracted['C_T_S_NAME'] = "N/A"
    
    # Add INSERT_USER using rule engine if not already set
    if 'INSERT_USER' not in extracted:
        rule_type, parser_path, insert_user = get_travel_agency_rule(
            extracted.get('C_T_S_NAME', c_t_s_name), sender_email, text
        )
        extracted['INSERT_USER'] = insert_user
    
    return extracted


def get_rule_based_search_folders(rule_type, outlook, namespace):
    """
    Get appropriate Outlook folders to search based on rule type
    Returns list of folder objects to search in
    """
    try:
        folders_to_search = []
        
        # Get the default inbox first
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        folders_to_search.append(inbox)
        
        # Add rule-specific folders based on rule type
        if rule_type.startswith("INNLINKWAY"):
            # Search for INNLINKWAY specific folders
            try:
                # Look for folders that might contain INNLINKWAY emails
                for folder in inbox.Folders:
                    folder_name = folder.Name.lower()
                    if any(keyword in folder_name for keyword in ['innlink', 'millennium', 'booking', 'reservation']):
                        folders_to_search.append(folder)
            except:
                pass
        
        elif rule_type.startswith("TRAVEL_AGENCY"):
            # Search for travel agency folders
            try:
                for folder in inbox.Folders:
                    folder_name = folder.Name.lower()
                    if any(keyword in folder_name for keyword in ['travel', 'agency', 'booking', 'tour']):
                        folders_to_search.append(folder)
            except:
                pass
        
        elif rule_type.startswith("AIRLINES"):
            # Search for airline folders
            try:
                for folder in inbox.Folders:
                    folder_name = folder.Name.lower()
                    if any(keyword in folder_name for keyword in ['airline', 'flight', 'air']):
                        folders_to_search.append(folder)
            except:
                pass
        
        # Always include sent items for outbound correspondence
        try:
            sent_items = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
            folders_to_search.append(sent_items)
        except:
            pass
            
        return folders_to_search
        
    except Exception as e:
        logger.error(f"Error getting rule-based search folders: {e}")
        # Fallback to just inbox
        try:
            return [namespace.GetDefaultFolder(6)]
        except:
            return []

def get_current_mailbox_info(outlook, namespace):
    """Get information about the current active mailbox"""
    try:
        # Get the active explorer (current Outlook window)
        explorer = outlook.ActiveExplorer()
        if explorer:
            current_folder = explorer.CurrentFolder()
            # Try to get the parent store (mailbox)
            store = current_folder.Store
            return current_folder, store
        else:
            return None, None
    except Exception as e:
        logger.warning(f"Could not get current mailbox info: {e}")
        return None, None

def search_all_folders_in_mailbox(store, guest_name, first_name="", days=2):
    """Search specific folders in the current mailbox for a specific guest
    Focus on: 2025\\Aug, 2025\\July, Groups, 0 OTA Notification, Inbox folders"""
    all_matching_emails = []
    
    def search_folder_recursive(folder, depth=0):
        nonlocal all_matching_emails
        
        try:
            folder_path = folder.FolderPath.lower()
            folder_name = folder.Name.lower()
            
            # Skip system folders that might cause issues
            if any(skip in folder_name for skip in ['calendar', 'contacts', 'tasks', 'notes', 'journal']):
                return
            
            # Check if this folder should be searched based on priority folders
            should_search = False
            
            # Priority folders: Inbox, Sent Items, Groups, 0 OTA Notification, and specific 2025 subfolders
            if ('inbox' in folder_name or 
                'sent items' in folder_name or 
                'groups' in folder_name or 
                '0 ota notification' in folder_name or
                ('2025' in folder_path and ('aug' in folder_path or 'july' in folder_path))):
                should_search = True
                logger.info(f"Searching priority folder: {folder.FolderPath}")
            elif depth == 0:  # Always search root level folders
                should_search = True
            
            # Get items in this folder if we should search it
            if should_search:
                items = folder.Items
                
                if len(items) > 0:
                    # Apply date filter (2 days)
                    since_date = (datetime.now() - timedelta(days=days)).strftime("%m/%d/%Y")
                    try:
                        filtered_items = items.Restrict(f'[ReceivedTime] >= "{since_date}" OR [SentOn] >= "{since_date}"')
                        
                        # Search through filtered items using both full name and first name
                        matches_in_folder = search_items_in_folder_for_guest(filtered_items, folder.Name, guest_name, first_name)
                        all_matching_emails.extend(matches_in_folder)
                        
                        if matches_in_folder:
                            logger.info(f"Found {len(matches_in_folder)} matches in {folder.FolderPath}")
                            
                    except Exception as e:
                        logger.warning(f"Could not filter items in folder {folder.Name}: {e}")
            
            # Search subfolders
            if folder.Folders.Count > 0:
                for subfolder in folder.Folders:
                    search_folder_recursive(subfolder, depth + 1)
                    
        except Exception as e:
            logger.warning(f"Error searching folder {folder.Name}: {e}")
    
    # Start recursive search from the root folder of the store
    try:
        root_folder = store.GetRootFolder()
        search_folder_recursive(root_folder)
    except Exception as e:
        logger.error(f"Could not access root folder: {e}")
    
    return all_matching_emails

def search_items_in_folder_for_guest(items, folder_name, guest_name, first_name=""):
    """Search for matching items in a specific folder for a guest using both full name and first name"""
    matching_emails = []
    
    for item in items:
        try:
            # Check if this is an email item
            if not hasattr(item, 'Subject'):
                continue
            
            # Get email properties
            sender_email = getattr(item, 'SenderEmailAddress', '') or ''
            sender_name = getattr(item, 'SenderName', '') or ''
            subject = getattr(item, 'Subject', '') or ''
            body = getattr(item, 'Body', '') or ''
            received_time = getattr(item, 'ReceivedTime', '') or getattr(item, 'SentOn', '')
            
            # Check if this email matches our guest criteria
            email_text = (subject + ' ' + body + ' ' + sender_email + ' ' + sender_name).lower()
            
            # Look for both full name and first name variations
            name_found = False
            
            # Search by full name (FULL_NAME column)
            if guest_name and guest_name.strip():
                guest_name_lower = guest_name.lower()
                name_parts = guest_name_lower.split()
                name_found = any(part in email_text for part in name_parts if len(part) > 2)
            
            # Search by first name (FIRST_NAME column)
            if not name_found and first_name and first_name.strip():
                first_name_lower = first_name.lower()
                if len(first_name_lower) > 2:
                    name_found = first_name_lower in email_text
            
            # Also check for specific senders (always include reservation emails)
            is_reservations_email = 'reservations.gmhd@millenniumhotels.com' in sender_email.lower()
            
            if name_found or is_reservations_email:
                email_info = {
                    'subject': subject,
                    'sender': sender_email,
                    'sender_name': sender_name,
                    'received_time': received_time,
                    'attachments': [],
                    'extracted_data': {},
                    'matched_reservation': guest_name,
                    'folder': folder_name
                }
                
                # For noreply-reservations emails, extract data from the email body and subject
                if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
                    # Combine subject and body for extraction
                    full_content = subject + "\n" + body
                    extracted_fields = extract_reservation_fields(full_content, sender_email)
                    email_info['extracted_data'] = extracted_fields
                    
                    # Format currency fields
                    for field in ['NET_TOTAL', 'TDF']:
                        if extracted_fields.get(field) != 'N/A' and extracted_fields.get(field):
                            try:
                                amount = float(str(extracted_fields[field]).replace(',', ''))
                                extracted_fields[f'{field}_AED'] = f"AED {amount:,.2f}"
                            except:
                                pass
                
                # Process PDF attachments if present
                if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                    for attachment in item.Attachments:
                        filename = getattr(attachment, 'FileName', '')
                        
                        if filename and filename.lower().endswith('.pdf'):
                            try:
                                # Save attachment temporarily with safe filename
                                safe_filename = f"temp_{filename.replace(' ', '_').replace('/', '_').replace('\\', '_')}"
                                temp_path = os.path.join(os.getcwd(), safe_filename)
                                
                                logger.info(f"Processing PDF attachment: {filename}")
                                attachment.SaveAsFile(temp_path)
                                
                                if os.path.exists(temp_path):
                                    with open(temp_path, 'rb') as f:
                                        pdf_data = f.read()
                                        logger.info(f"PDF size: {len(pdf_data)} bytes")
                                        text = extract_pdf_text(pdf_data)
                                        
                                        if text and len(text.strip()) > 10:  # Minimum text threshold
                                            logger.info(f"Extracted {len(text)} characters from PDF")
                                            extracted_fields = extract_reservation_fields(text, sender_email)
                                            
                                            # Format currency fields including NET
                                            for field in ['NET', 'NET_TOTAL', 'TDF', 'AMOUNT', 'TOTAL']:
                                                if extracted_fields.get(field) != 'N/A' and extracted_fields.get(field):
                                                    try:
                                                        amount_str = str(extracted_fields[field]).replace(',', '')
                                                        amount = float(amount_str)
                                                        extracted_fields[f'{field}_AED'] = f"AED {amount:,.2f}"
                                                    except ValueError:
                                                        logger.warning(f"Could not parse currency field {field}: {extracted_fields[field]}")
                                            
                                            email_info['extracted_data'] = extracted_fields
                                            email_info['attachments'].append({
                                                'filename': filename,
                                                'size': len(pdf_data),
                                                'text_extracted': True,
                                                'text_length': len(text),
                                                'contains_china_southern': 'china southern' in text.lower()
                                            })
                                            logger.info(f"Successfully processed PDF: {filename}")
                                        else:
                                            logger.warning(f"Insufficient text extracted from PDF: {filename}")
                                            email_info['attachments'].append({
                                                'filename': filename,
                                                'size': len(pdf_data),
                                                'text_extracted': False,
                                                'error': 'No readable text found'
                                            })
                                else:
                                    logger.error(f"Failed to save PDF attachment: {filename}")
                                
                                # Clean up temp file
                                if os.path.exists(temp_path):
                                    try:
                                        os.remove(temp_path)
                                    except Exception as cleanup_error:
                                        logger.warning(f"Could not clean up temp file {temp_path}: {cleanup_error}")
                                    
                            except Exception as e:
                                logger.warning(f"Error processing PDF {filename}: {e}")
                        else:
                            email_info['attachments'].append({
                                'filename': filename,
                                'type': 'non-pdf'
                            })
                
                matching_emails.append(email_info)
                
        except Exception as e:
            continue  # Skip problematic items
    
    return matching_emails

def search_emails_for_reservation(outlook, namespace, reservation_data, days=2):
    """Search emails for a specific reservation using guest name and dates - Enhanced with current mailbox search"""
    try:
        # Create search criteria from Entered On sheet columns
        guest_name = reservation_data.get('FULL_NAME', '').strip()  # Column for full name
        first_name = reservation_data.get('FIRST_NAME', '').strip()  # Column for first name
        arrival_date = reservation_data.get('ARRIVAL')
        
        if not guest_name and not first_name:
            return []
        
        # Get current mailbox info
        current_folder, store = get_current_mailbox_info(outlook, namespace)
        
        if not store:
            # Fallback to default inbox if we can't get current mailbox
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            store = inbox.Store
        
        # Search specific folders in the current mailbox (2025\Aug, 2025\July, Groups, Inbox)
        matching_emails = search_all_folders_in_mailbox(store, guest_name, first_name, days)
        
        return matching_emails
        
    except Exception as e:
        logger.error(f"Error searching emails for {guest_name}: {e}")
        return []

def process_all_reservations_with_emails(outlook, namespace, reservations_df, days=7, run_id=None, db=None):
    """Process all reservations and search for matching emails"""
    results = []
    start_time = datetime.now()
    
    for idx, reservation in reservations_df.iterrows():
        reservation_dict = reservation.to_dict()
        
        # Search for emails related to this reservation
        matching_emails = search_emails_for_reservation(outlook, namespace, reservation_dict, days)
        
        # Combine reservation data with email findings
        result = {
            'reservation_index': idx,
            'reservation_data': reservation_dict,
            'matching_emails': matching_emails,
            'email_count': len(matching_emails),
            'has_pdf_data': any(email.get('extracted_data') for email in matching_emails),
            'status': 'EMAIL_FOUND' if matching_emails else 'NO_EMAIL_FOUND'
        }
        
        # If we found email data, merge it with reservation data
        if matching_emails:
            for email in matching_emails:
                if email.get('extracted_data'):
                    # Merge email extracted data with reservation data
                    for field, value in email['extracted_data'].items():
                        if value != 'N/A':
                            result['reservation_data'][f'MAIL_{field}'] = value
                    
                    # Also extract from body text using rule engine
                    sender_email = getattr(email, 'sender', '')
                    c_t_s_name = result['reservation_data'].get('C_T_S_NAME', '')
                    
                    # Get subject and body content
                    email_text = f"{email.get('subject', '')}\n{getattr(email, 'body', '')}"
                    
                    # Use rule engine to get INSERT_USER
                    rule_type, parser_path, insert_user = get_travel_agency_rule(c_t_s_name, sender_email, email_text)
                    
                    # Extract additional fields using rule-based extraction
                    additional_fields = extract_reservation_fields(email_text, sender_email, c_t_s_name)
                    
                    # Add INSERT_USER to the extracted data
                    additional_fields['INSERT_USER'] = insert_user
                    
                    for field, value in additional_fields.items():
                        if value != 'N/A':
                            result['reservation_data'][f'MAIL_{field}'] = value
        
        results.append(result)
        
        # Add progress feedback
        if idx % 10 == 0:
            logger.info(f"Processed {idx + 1}/{len(reservations_df)} reservations")
    
    # Save email extraction results to database
    if run_id and db:
        try:
            saved_count = db.save_email_extraction(results, run_id)
            execution_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"Saved {saved_count} email extractions to database in {execution_time:.2f}s")
        except Exception as e:
            logger.error(f"Failed to save email extraction to database: {e}")
            if db:
                db.log_error(run_id, str(e), "process_all_reservations_with_emails")
    
    return results

def perform_audit_checks(df, email_data=None, run_id=None, db=None):
    """Perform audit validation checks on the data including email extraction comparison"""
    if df is None or df.empty:
        return pd.DataFrame()
    
    start_time = datetime.now()
    
    df_audit = df.copy()
    df_audit['audit_status'] = 'PASS'
    df_audit['audit_issues'] = ''
    df_audit['fields_matching'] = 0
    df_audit['total_email_fields'] = 0
    df_audit['match_percentage'] = 0
    df_audit['email_vs_data_status'] = 'N/A'
    
    # Initialize Mail_ columns with N/A
    mail_fields = ['FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 
                 'RATE_CODE', 'C_T_S', 'C_T_S_NAME', 'NET', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
    
    for field in mail_fields:
        df_audit[f'Mail_{field}'] = 'N/A'
    
    # Create email data lookup for comparison
    email_lookup = {}
    if email_data:
        for result in email_data:
            guest_name = result['reservation_data'].get('FULL_NAME', '')
            email_lookup[guest_name] = result
    
    for idx, row in df_audit.iterrows():
        row_issues = []
        
        # Check 1: NIGHTS = Departure - Arrival (using dd/mm/yyyy format)
        if pd.notna(row['ARRIVAL']) and pd.notna(row['DEPARTURE']):
            try:
                # Always use dayfirst=True for dd/mm/yyyy format
                arrival = pd.to_datetime(row['ARRIVAL'], dayfirst=True)
                departure = pd.to_datetime(row['DEPARTURE'], dayfirst=True)
                calculated_nights = (departure - arrival).days
                
                if pd.notna(row['NIGHTS']) and abs(row['NIGHTS'] - calculated_nights) > 0:
                    row_issues.append(f"Night calculation mismatch: Expected {calculated_nights}, got {row['NIGHTS']}")
            except:
                row_issues.append("Invalid date format (expected dd/mm/yyyy)")
        
        # Check 2: NET_TOTAL >= TDF (if both exist)
        if pd.notna(row.get('NET_TOTAL')) and pd.notna(row.get('TDF')):
            try:
                net_total = float(str(row['NET_TOTAL']).replace(',', ''))
                tdf = float(str(row['TDF']).replace(',', ''))
                if net_total < tdf:
                    row_issues.append(f"NET_TOTAL ({net_total}) < TDF ({tdf})")
            except:
                row_issues.append("Invalid numeric format for NET_TOTAL or TDF")
        
        # Check 3: PERSONS > 0
        if pd.notna(row.get('PERSONS')):
            try:
                persons = int(row['PERSONS'])
                if persons <= 0:
                    row_issues.append(f"Invalid person count: {persons}")
            except:
                row_issues.append("Invalid person count format")
        
        # Check 4: Required fields present
        required_fields = ['FULL_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS']
        rate_fields = ['NET_TOTAL', 'ROOM_RATE', 'ADR', 'TOTAL_AMOUNT']
        
        for field in required_fields:
            if pd.isna(row.get(field)) or row.get(field) == '' or row.get(field) == 'N/A':
                row_issues.append(f"Missing required field: {field}")
        
        # Check 5: At least one rate field should be present
        has_rate_info = any(pd.notna(row.get(f'MAIL_{field}')) or pd.notna(row.get(field)) 
                           for field in rate_fields)
        if not has_rate_info:
            row_issues.append("Missing rate information - no rate fields found")
        
        # NEW: Check 6: Email extraction vs converted data comparison
        guest_name = row.get('FULL_NAME', '')
        if guest_name in email_lookup:
            email_result = email_lookup[guest_name]
            email_fields = {}
            
            # Gather all email extracted fields
            for email in email_result.get('matching_emails', []):
                if email.get('extracted_data'):
                    email_fields.update(email['extracted_data'])
            
            # ADD MAIL_ COLUMNS TO AUDIT DATAFRAME
            mail_fields = ['FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 
                         'RATE_CODE', 'C_T_S', 'C_T_S_NAME', 'NET', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
            
            for field in mail_fields:
                mail_col = f'Mail_{field}'
                df_audit.at[idx, mail_col] = email_fields.get(field, 'N/A')
            
            # Compare fields between email extraction and converted data
            comparison_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM']
            matching_fields = 0
            total_comparable_fields = 0
            
            for field in comparison_fields:
                email_value = email_fields.get(field, 'N/A')
                data_value = str(row.get(field, 'N/A'))
                
                if email_value != 'N/A' and data_value != 'N/A':
                    total_comparable_fields += 1
                    
                    # Normalize for comparison
                    if field in ['ARRIVAL', 'DEPARTURE']:
                        try:
                            # Always use dayfirst=True for dd/mm/yyyy format
                            email_date = pd.to_datetime(email_value, dayfirst=True).strftime('%d/%m/%Y')
                            data_date = pd.to_datetime(data_value, dayfirst=True).strftime('%d/%m/%Y')
                            if email_date == data_date:
                                matching_fields += 1
                        except:
                            pass  # Date format mismatch
                    elif field == 'ROOM':
                        # Special handling for ST and SK room group - they're equivalent (twin/king bed)
                        email_room = str(email_value).lower().strip()
                        data_room = str(data_value).lower().strip()
                        if (email_room == data_room or 
                            (email_room in ['st', 'sk'] and data_room in ['st', 'sk'])):
                            matching_fields += 1
                    elif str(email_value).lower().strip() == str(data_value).lower().strip():
                        matching_fields += 1
            
            df_audit.at[idx, 'fields_matching'] = matching_fields
            df_audit.at[idx, 'total_email_fields'] = total_comparable_fields
            
            if total_comparable_fields > 0:
                match_percentage = (matching_fields / total_comparable_fields) * 100
                df_audit.at[idx, 'match_percentage'] = match_percentage
                
                if match_percentage >= 80:
                    df_audit.at[idx, 'email_vs_data_status'] = 'PASS'
                elif match_percentage >= 60:
                    df_audit.at[idx, 'email_vs_data_status'] = 'WARNING'
                else:
                    df_audit.at[idx, 'email_vs_data_status'] = 'FAIL'
                    row_issues.append(f"Low email-data match: {match_percentage:.1f}% ({matching_fields}/{total_comparable_fields} fields)")
            else:
                df_audit.at[idx, 'email_vs_data_status'] = 'NO_EMAIL_DATA'
        
        # Update audit status
        if row_issues:
            df_audit.at[idx, 'audit_status'] = 'FAIL'
            df_audit.at[idx, 'audit_issues'] = '; '.join(row_issues)
    
    # Save audit results to database
    if run_id and db:
        try:
            saved_count = db.save_audit_results(df_audit, run_id)
            execution_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"Saved {saved_count} audit results to database in {execution_time:.2f}s")
        except Exception as e:
            logger.error(f"Failed to save audit results to database: {e}")
            if db:
                db.log_error(run_id, str(e), "perform_audit_checks")
    
    return df_audit

# Streamlit App
def get_latest_file_from_path(base_path="P:\\Reservation\\Entered on"):
    """Get the latest .xlsm file from the latest folder in the specified path"""
    try:
        if not os.path.exists(base_path):
            return None, f"Base path does not exist: {base_path}"
        
        # Get all directories in the base path
        directories = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
        
        if not directories:
            return None, "No directories found in base path"
        
        # Sort directories by modification time (latest first)
        directories.sort(key=lambda x: os.path.getmtime(os.path.join(base_path, x)), reverse=True)
        latest_dir = directories[0]
        latest_dir_path = os.path.join(base_path, latest_dir)
        
        # Get all .xlsm files in the latest directory, skip temporary files
        xlsm_files = [f for f in os.listdir(latest_dir_path) 
                     if f.lower().endswith('.xlsm') and not f.startswith('~$')]
        
        if not xlsm_files:
            return None, f"No .xlsm files found in latest directory: {latest_dir}"
        
        # Sort files by modification time (latest first)
        xlsm_files.sort(key=lambda x: os.path.getmtime(os.path.join(latest_dir_path, x)), reverse=True)
        latest_file = xlsm_files[0]
        latest_file_path = os.path.join(latest_dir_path, latest_file)
        
        return latest_file_path, f"Selected: {latest_dir}\\{latest_file}"
        
    except Exception as e:
        return None, f"Error finding latest file: {e}"

def main():
    st.title("🏨 Entered On Audit System")
    st.markdown("---")
    
    # Quick info about improvements
    with st.sidebar.expander("🆕 Recent Improvements"):
        st.write("• Complete SQLite integration with persistent storage")
        st.write("• Enhanced Logs & History tab with run tracking")
        st.write("• Improved UI with collapsible sections")
        st.write("• Enhanced current mailbox search (all folders)")
        st.write("• Added 0 OTA Notification folder search")
        st.write("• Support for noreply-reservations@millenniumhotels.com emails")
        st.write("• Added comprehensive rate extraction (ADR, TDF calculation)")
        st.write("• Currency set to AED only")
        st.write("• Email vs data field matching audit")
        st.write("• **NEW**: Travel Agency TO parsers (Travco, Dubai Link, Dakkak, Duri)")
        
        # Add parser information
        with st.expander("🔧 Available Email Parsers", expanded=False):
            st.markdown("### Specialized Travel Agency Parsers")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**Travco Parser**")
                st.write("• Email: travco@travco.co.uk")
                st.write("• Detects: Hotel Booking Confirmation")
                st.write("• Rate Code: TOBBJN (TO* format)")
                st.write("• Fields: All standard extraction fields")
                st.success("✅ Active")
            
            with col2:
                st.markdown("**Dubai Link Parser**")
                st.write("• Email: gte.travel domain")
                st.write("• Detects: Confirmed Booking")
                st.write("• Rate Code: TO* format")
                st.write("• Fields: All standard extraction fields")
                st.success("✅ Active")
            
            with col3:
                st.markdown("**Dakkak Parser**")
                st.write("• Email: dakkak.com (Dakkak DMC)")
                st.write("• Detects: Hotel New Booking")
                st.write("• Rate Code: BKGHO format")
                st.write("• Fields: All standard extraction fields")
                st.success("✅ Active")
            
            with col4:
                st.markdown("**Duri Parser**")
                st.write("• Email: hanmail.net (Duri)")
                st.write("• Detects: DURI TRAVEL booking")
                st.write("• Rate Code: AED rate format")
                st.write("• Fields: All standard extraction fields")
                st.success("✅ Active")
                
            st.markdown("### Standard Pattern Parsers")
            st.write("• noreply-reservations@millenniumhotels.com (INNLINKWAY)")
            st.write("• China Southern Air reservations")
            st.write("• Generic hotel confirmation emails")
        st.write("• Automatic file selection from P:\\Reservation\\Entered on")
    
    # Hardcode days to 2
    email_days = 2
    
    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📧 Email Extraction Results", "📊 Converted Data", "🔍 Audit Results", "📝 Logs & History"])
    
    # Tab 1: Email Extraction Results
    with tab1:
        st.header("📧 Email Extraction Results")
        
        # Email search controls
        with st.expander("🔧 Email Search Configuration", expanded=False):
            st.info("📧 **Outlook Integration**: Connects to your local Outlook installation")
            st.info("🗓️ **Search Period**: Last 2 days automatically")
            st.info("📂 **Search Folders**: 2025\\Aug, 2025\\July, Groups, 0 OTA Notification, Inbox, Sent Items")
            st.info("💱 **Currency**: All amounts displayed in AED only")
        
        # Email processing button
        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button("🔄 Search Emails for Each Reservation", type="primary"):
                if st.session_state.processed_data is not None:
                    with st.spinner("Connecting to Outlook and searching emails..."):
                        try:
                            outlook, namespace = connect_to_outlook()
                            if outlook and namespace:
                                # Process all reservations and search for emails
                                email_results = process_all_reservations_with_emails(
                                    outlook, namespace, st.session_state.processed_data, email_days,
                                    run_id=st.session_state.current_run_id, db=st.session_state.database
                                )
                                st.session_state.email_data = email_results
                                
                                # Summary stats
                                total_reservations = len(email_results)
                                with_emails = sum(1 for r in email_results if r['email_count'] > 0)
                                with_pdf_data = sum(1 for r in email_results if r['has_pdf_data'])
                                
                                st.success(f"✅ Processed {total_reservations} reservations")
                                st.success(f"📧 Found emails for {with_emails} reservations")
                                st.success(f"📄 Extracted PDF data for {with_pdf_data} reservations")
                            else:
                                st.error("❌ Could not connect to Outlook")
                        except Exception as e:
                            st.error(f"Error processing emails: {e}")
                else:
                    st.warning("Please upload an Excel file first")
        
        with col2:
            if st.session_state.email_data:
                st.metric("Email Status", "✅ Complete")
            else:
                st.metric("Email Status", "⏳ Pending")
        
        st.markdown("---")
        
        # File selection section
        with st.expander("📂 Data Source Selection", expanded=True):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # Auto-load on first run
                if not st.session_state.auto_loaded:
                    latest_file, status_msg = get_latest_file_from_path()
                    if latest_file:
                        st.session_state.selected_file_path = latest_file
                        try:
                            result = process_entered_on_report(latest_file)
                            if len(result) == 3:  # With database (DataFrame, csv_path, run_id)
                                processed_df, csv_path, run_id = result
                                st.session_state.current_run_id = run_id
                            else:  # Without database (DataFrame, csv_path)
                                processed_df, csv_path = result
                            
                            st.session_state.processed_data = processed_df
                            st.session_state.uploaded_file_name = os.path.basename(latest_file)
                            st.session_state.auto_loaded = True
                            st.success(f"✅ Auto-loaded: {status_msg} ({len(processed_df)} records)")
                        except Exception as e:
                            st.warning(f"Auto-load failed: {e}")
                
                # Manual refresh button
                if st.button("🔄 Refresh - Select Latest File from P:\\Reservation\\Entered on"):
                    latest_file, status_msg = get_latest_file_from_path()
                    if latest_file:
                        st.session_state.selected_file_path = latest_file
                        st.success(status_msg)
                    
                        # Auto-convert the file
                        try:
                            with st.spinner("Auto-processing Excel file..."):
                                result = process_entered_on_report(latest_file)
                                if len(result) == 3:  # With database (DataFrame, csv_path, run_id)
                                    processed_df, csv_path, run_id = result
                                    st.session_state.current_run_id = run_id
                                else:  # Without database (DataFrame, csv_path)
                                    processed_df, csv_path = result
                                
                                st.session_state.processed_data = processed_df
                                st.session_state.uploaded_file_name = os.path.basename(latest_file)
                            st.success(f"✅ Auto-processed {len(processed_df)} records")
                        except Exception as e:
                            st.error(f"Error auto-processing file: {e}")
                    else:
                        st.error(status_msg)
        
            with col2:
                # Manual file upload as fallback
                uploaded_file = st.file_uploader(
                    "Or manually upload Excel file", 
                    type=['xlsm', 'xlsx'],
                    help="Select the Entered On report Excel file"
                )
                
                if uploaded_file is not None:
                    if st.session_state.uploaded_file_name != uploaded_file.name:
                        st.session_state.uploaded_file_name = uploaded_file.name
                        
                        # Save uploaded file temporarily
                        temp_file_path = f"temp_{uploaded_file.name}"
                        with open(temp_file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    
                        try:
                            # Process the Excel file
                            with st.spinner("Processing Excel file..."):
                                result = process_entered_on_report(temp_file_path)
                                if len(result) == 3:  # With database (DataFrame, csv_path, run_id)
                                    processed_df, csv_path, run_id = result
                                    st.session_state.current_run_id = run_id
                                else:  # Without database (DataFrame, csv_path)
                                    processed_df, csv_path = result
                                
                                st.session_state.processed_data = processed_df
                            st.success(f"✅ Processed {len(processed_df)} records")
                        except Exception as e:
                            st.error(f"Error processing file: {e}")
                        finally:
                            # Clean up temp file
                            if os.path.exists(temp_file_path):
                                os.remove(temp_file_path)
            
            # Show current file status
            if st.session_state.processed_data is not None:
                st.info(f"📄 Currently loaded: {st.session_state.uploaded_file_name} ({len(st.session_state.processed_data)} records)")
            else:
                st.warning("📄 No file loaded. Please use refresh button or manual upload above.")
        
        st.markdown("---")
        
        if st.session_state.email_data:
            email_results = st.session_state.email_data
            
            # Summary metrics
            with st.expander("📊 Email Search Summary", expanded=True):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Reservations", len(email_results))
                with col2:
                    reservations_with_emails = sum(1 for r in email_results if r['email_count'] > 0)
                    st.metric("Found Emails", reservations_with_emails)
                with col3:
                    reservations_with_data = sum(1 for r in email_results if r['has_pdf_data'])
                    st.metric("PDF Data Extracted", reservations_with_data)
                with col4:
                    total_emails = sum(r['email_count'] for r in email_results)
                    st.metric("Total Emails", total_emails)
            
            # Filter options
            with st.expander("🔍 Email Result Filters", expanded=False):
                status_filter = st.selectbox("Filter by Status", ["All", "EMAIL_FOUND", "NO_EMAIL_FOUND"])
                filtered_results = email_results
                if status_filter != "All":
                    filtered_results = [r for r in email_results if r['status'] == status_filter]
            
            # Email Extraction Results Table - Mail extraction variables only
            st.subheader("📄 Email Extraction Results")
            
            # Create simplified table showing only MAIL extraction variables
            table_data = []
            
            # Core mail extraction fields to display
            mail_fields = ['FIRST_NAME', 'FULL_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 
                          'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
            
            for result in filtered_results:
                reservation = result['reservation_data']
                guest_name = reservation.get('FULL_NAME', 'N/A')
                
                # Get all extracted email data for this guest
                email_data = {}
                for email in result.get('matching_emails', []):
                    if email.get('extracted_data'):
                        email_data.update(email['extracted_data'])
                
                # Create row with guest name and mail extraction fields only
                row_data = {
                    'Guest_Name': guest_name,
                    'Email_Status': result['status'],
                    'Emails_Found': result['email_count']
                }
                
                # Add mail extraction fields with MAIL_ prefix
                for field in mail_fields:
                    mail_value = email_data.get(field, 'N/A')
                    # Format currency fields
                    if field in ['NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT'] and mail_value != 'N/A':
                        try:
                            amount = float(str(mail_value).replace(',', ''))
                            mail_value = f"AED {amount:,.2f}"
                        except:
                            mail_value = 'N/A'
                    row_data[f'MAIL_{field}'] = mail_value
                
                table_data.append(row_data)
            
            if table_data:
                results_df = pd.DataFrame(table_data)
                st.dataframe(results_df, use_container_width=True, height=500)
                
                # Save to database button
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("💾 Save Email Extractions to Database", type="secondary"):
                        if st.session_state.current_run_id:
                            try:
                                # Save current email extraction results to database
                                saved_count = st.session_state.database.save_email_extraction(
                                    st.session_state.email_data, 
                                    st.session_state.current_run_id
                                )
                                st.success(f"✅ Saved {saved_count} email extractions to database")
                            except Exception as e:
                                st.error(f"❌ Error saving to database: {e}")
                        else:
                            st.warning("⚠️ No active run ID. Please process data first.")
            else:
                st.info("No results to display.")
            
            # Section to view previously processed extractions
            st.markdown("---")
            st.subheader("📚 Previously Processed Email Extractions")
            
            # Get recent runs for selection
            try:
                recent_runs = st.session_state.database.get_recent_runs(limit=10)
                
                if not recent_runs.empty:
                    # Select run to view
                    run_options = recent_runs['run_id'].tolist()
                    selected_run = st.selectbox(
                        "Select a run to view email extractions:",
                        options=[''] + run_options,
                        format_func=lambda x: f"Current Run" if x == st.session_state.current_run_id else (f"{x[-8:]} - {recent_runs[recent_runs['run_id']==x]['run_timestamp'].iloc[0]}" if x else "Select a run...")
                    )
                    
                    if selected_run:
                        # Load email extraction data for selected run
                        email_extraction_df = st.session_state.database.export_data('reservations_email', selected_run)
                        
                        if not email_extraction_df.empty:
                            # Transform database data to display format
                            display_data = []
                            for _, row in email_extraction_df.iterrows():
                                display_row = {
                                    'Guest_Name': row['guest_name'],
                                    'Email_Subject': row['email_subject'][:50] + '...' if len(row['email_subject']) > 50 else row['email_subject'],
                                    'Email_Sender': row['email_sender'],
                                    'Folder': row['folder_name'],
                                    'MAIL_FIRST_NAME': row['mail_first_name'],
                                    'MAIL_ARRIVAL': row['mail_arrival'],
                                    'MAIL_DEPARTURE': row['mail_departure'],
                                    'MAIL_NIGHTS': row['mail_nights'],
                                    'MAIL_PERSONS': row['mail_persons'],
                                    'MAIL_ROOM': row['mail_room'],
                                    'MAIL_RATE_CODE': row['mail_rate_code'],
                                    'MAIL_C_T_S': row['mail_c_t_s'],
                                    'MAIL_NET_TOTAL': f"AED {row['mail_net_total']:,.2f}" if row['mail_net_total'] > 0 else 'N/A',
                                    'MAIL_TOTAL': f"AED {row['mail_total']:,.2f}" if row['mail_total'] > 0 else 'N/A',
                                    'MAIL_TDF': f"AED {row['mail_tdf']:,.2f}" if row['mail_tdf'] > 0 else 'N/A',
                                    'MAIL_ADR': f"AED {row['mail_adr']:,.2f}" if row['mail_adr'] > 0 else 'N/A',
                                    'MAIL_AMOUNT': f"AED {row['mail_amount']:,.2f}" if row['mail_amount'] > 0 else 'N/A'
                                }
                                display_data.append(display_row)
                            
                            previous_df = pd.DataFrame(display_data)
                            
                            # Show summary
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Total Email Extractions", len(previous_df))
                            with col2:
                                unique_guests = previous_df['Guest_Name'].nunique()
                                st.metric("Unique Guests", unique_guests)
                            with col3:
                                with_data = len([row for row in display_data if any(v != 'N/A' for k, v in row.items() if k.startswith('MAIL_'))])
                                st.metric("With Extraction Data", with_data)
                            
                            # Display the data
                            st.dataframe(previous_df, use_container_width=True, height=400)
                            
                            # Export option
                            csv_data = previous_df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download Previous Email Extractions CSV",
                                data=csv_data,
                                file_name=f"previous_email_extractions_{selected_run}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv"
                            )
                        else:
                            st.info("No email extraction data found for this run.")
                else:
                    st.info("No previous runs found in the database.")
                    
            except Exception as e:
                st.error(f"❌ Error loading previous extractions: {e}")
            
            # Export results
            st.markdown("---")
            if st.button("📥 Export Email Results"):
                # Create export DataFrame
                export_data = []
                for result in email_results:
                    reservation = result['reservation_data']
                    base_row = {
                        'Guest_Name': reservation.get('FULL_NAME', ''),
                        'Arrival': reservation.get('ARRIVAL', ''),
                        'Departure': reservation.get('DEPARTURE', ''),
                        'Nights': reservation.get('NIGHTS', ''),
                        'Room': reservation.get('ROOM', ''),
                        'Amount_AED': reservation.get('AMOUNT', ''),
                        'Email_Status': result['status'],
                        'Emails_Found': result['email_count'],
                        'PDF_Data_Found': result['has_pdf_data']
                    }
                    
                    # Add email extracted fields
                    for field, value in reservation.items():
                        if field.startswith('EMAIL_'):
                            base_row[field] = value
                    
                    export_data.append(base_row)
                
                export_df = pd.DataFrame(export_data)
                csv = export_df.to_csv(index=False)
                st.download_button(
                    label="💾 Download Email Results CSV",
                    data=csv,
                    file_name=f"email_extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        else:
            st.info("👆 Load an Excel file using the options above, then click 'Search Emails for Each Reservation' button.")
    
    # Tab 2: Converted Data (Full Entered On sheet)
    with tab2:
        st.header("📊 Converted Data - Full Entered On Sheet")
        
        if st.session_state.processed_data is not None:
            df = st.session_state.processed_data
            
            # Summary metrics
            with st.expander("📊 Summary Statistics", expanded=True):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Records", len(df))
                with col2:
                    total_amount = df['AMOUNT'].sum() if 'AMOUNT' in df.columns else 0
                    st.metric("Total Amount (AED)", f"AED {total_amount:,.2f}")
                with col3:
                    total_nights = df['NIGHTS'].sum() if 'NIGHTS' in df.columns else 0
                    st.metric("Total Nights", f"{total_nights:,}")
                with col4:
                    avg_adr = df['ADR'].mean() if 'ADR' in df.columns else 0
                    st.metric("Average ADR (AED)", f"AED {avg_adr:.2f}")
            
            # Filters
            with st.expander("🔍 Data Filters", expanded=False):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if 'SEASON' in df.columns:
                        seasons = ['All'] + list(df['SEASON'].unique())
                        selected_season = st.selectbox("Season", seasons)
                        if selected_season != 'All':
                            df = df[df['SEASON'] == selected_season]
                
                with col2:
                    if 'COMPANY_CLEAN' in df.columns:
                        companies = ['All'] + list(df['COMPANY_CLEAN'].unique())
                        selected_company = st.selectbox("Company", companies)
                        if selected_company != 'All':
                            df = df[df['COMPANY_CLEAN'] == selected_company]
                
                with col3:
                    if 'ROOM' in df.columns:
                        rooms = ['All'] + list(df['ROOM'].unique())
                        selected_room = st.selectbox("Room Type", rooms)
                        if selected_room != 'All':
                            df = df[df['ROOM'] == selected_room]
            
            # Display the full data
            st.subheader("📋 Full Dataset")
            st.dataframe(
                df,
                use_container_width=True,
                height=600
            )
            
            # Download button
            csv = df.to_csv(index=False)
            st.download_button(
                label="💾 Download as CSV",
                data=csv,
                file_name=f"entered_on_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            
        else:
            st.info("👆 Go to Email Extraction Results tab to load an Excel file first.")
    
    # Tab 3: Audit Results
    with tab3:
        st.header("🔍 Audit Results")
        
        if st.session_state.processed_data is not None:
            # Run audit button
            if st.button("🔄 Run Audit Checks"):
                with st.spinner("Performing audit checks including email extraction comparison..."):
                    audit_df = perform_audit_checks(
                        st.session_state.processed_data, st.session_state.email_data,
                        run_id=st.session_state.current_run_id, db=st.session_state.database
                    )
                    st.session_state.audit_results = audit_df
            
            if st.session_state.audit_results is not None:
                audit_df = st.session_state.audit_results
                
                # Audit summary - Enhanced with email extraction metrics
                with st.expander("📊 Audit Summary", expanded=True):
                    col1, col2, col3, col4, col5, col6 = st.columns(6)
                    with col1:
                        st.metric("Total Records", len(audit_df))
                    with col2:
                        pass_count = len(audit_df[audit_df['audit_status'] == 'PASS'])
                        st.metric("Passed", pass_count, delta=f"{pass_count/len(audit_df)*100:.1f}%")
                    with col3:
                        fail_count = len(audit_df[audit_df['audit_status'] == 'FAIL'])
                        st.metric("Failed", fail_count, delta=f"{fail_count/len(audit_df)*100:.1f}%")
                    with col4:
                        completion_rate = pass_count / len(audit_df) * 100
                        st.metric("Success Rate", f"{completion_rate:.1f}%")
                    with col5:
                        email_pass_count = len(audit_df[audit_df['email_vs_data_status'] == 'PASS'])
                        st.metric("Email Match PASS", email_pass_count)
                    with col6:
                        avg_match = audit_df['match_percentage'].mean()
                        st.metric("Avg Match %", f"{avg_match:.1f}%")
                
                # Enhanced filters
                with st.expander("🔍 Audit Filters", expanded=False):
                    col1, col2 = st.columns(2)
                    with col1:
                        status_filter = st.selectbox("Filter by Audit Status", ["All", "PASS", "FAIL"])
                    with col2:
                        email_filter = st.selectbox("Filter by Email Match", ["All", "PASS", "WARNING", "FAIL", "NO_EMAIL_DATA"])
                
                display_df = audit_df
                if status_filter != "All":
                    display_df = display_df[display_df['audit_status'] == status_filter]
                if email_filter != "All":
                    display_df = display_df[display_df['email_vs_data_status'] == email_filter]
                
                # Display audit results
                st.subheader("📊 Audit Results")
                
                # Configure columns to display - Side by side pairs: FIELD then Mail_FIELD
                # ARRIVAL, Mail_ARRIVAL, DEPARTURE, Mail_DEPARTURE, etc.
                display_columns = []
                
                # Start with name columns - side by side pairs
                display_columns.extend(['FULL_NAME', 'FIRST_NAME', 'Mail_FIRST_NAME'])
                
                # Create immediate side-by-side pairs: FIELD, Mail_FIELD
                comparison_fields = ['ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 
                                   'RATE_CODE', 'C_T_S', 'NET', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
                
                for field in comparison_fields:
                    # Add original field followed immediately by its Mail_ counterpart
                    display_columns.extend([field, f'Mail_{field}'])
                
                # Add audit result columns at the end
                audit_columns = ['fields_matching', 'total_email_fields', 'match_percentage', 
                               'email_vs_data_status', 'audit_status', 'audit_issues']
                display_columns.extend(audit_columns)
                available_columns = [col for col in display_columns if col in display_df.columns]
                
                # Apply conditional formatting to highlight mismatched Mail_ columns
                def highlight_mismatched_data(row):
                    styles = [''] * len(row)
                    
                    # Define comparison fields locally
                    comparison_fields_local = ['FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 
                                             'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
                    
                    # Compare each field with its Mail_ counterpart
                    for field in comparison_fields_local:
                        original_col = field
                        mail_col = f'Mail_{field}'
                        
                        if original_col in row.index and mail_col in row.index:
                            original_val = str(row[original_col]).strip() if pd.notna(row[original_col]) else 'N/A'
                            mail_val = str(row[mail_col]).strip() if pd.notna(row[mail_col]) else 'N/A'
                            
                            # Skip comparison if either is N/A
                            if original_val != 'N/A' and mail_val != 'N/A':
                                try:
                                    mail_col_idx = row.index.get_loc(mail_col)
                                    
                                    # For numeric fields (amounts), apply rounding to 2 decimal places and tolerance logic
                                    if field in ['NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']:
                                        try:
                                            # Clean and convert to float, removing AED prefix and commas
                                            original_clean = original_val.replace('AED', '').replace(',', '').strip()
                                            mail_clean = mail_val.replace('AED', '').replace(',', '').strip()
                                            
                                            original_num = round(float(original_clean), 2)
                                            mail_num = round(float(mail_clean), 2)
                                            
                                            # Calculate difference
                                            difference = abs(original_num - mail_num)
                                            
                                            if difference == 0:
                                                # Perfect match - mark green
                                                styles[mail_col_idx] = 'color: green; font-weight: bold'
                                            elif difference <= 1:
                                                # Within ±1 tolerance - don't mark red (no special color)
                                                pass
                                            else:
                                                # Outside tolerance - mark red
                                                styles[mail_col_idx] = 'color: red; font-weight: bold'
                                        except (ValueError, TypeError):
                                            # Non-numeric comparison fallback
                                            if original_val == mail_val:
                                                styles[mail_col_idx] = 'color: green; font-weight: bold'
                                            else:
                                                styles[mail_col_idx] = 'color: red; font-weight: bold'
                                    else:
                                        # Non-numeric fields - exact match comparison
                                        if original_val == mail_val:
                                            styles[mail_col_idx] = 'color: green; font-weight: bold'
                                        else:
                                            styles[mail_col_idx] = 'color: red; font-weight: bold'
                                except KeyError:
                                    continue
                    
                    return styles
                
                # Create styled dataframe with conditional formatting
                try:
                    styled_df = display_df[available_columns].style.apply(highlight_mismatched_data, axis=1)
                    st.dataframe(
                        styled_df,
                        use_container_width=True,
                        height=600
                    )
                except Exception as e:
                    # Fallback to regular dataframe if styling fails
                    st.warning(f"Conditional formatting failed: {e}")
                    st.dataframe(
                        display_df[available_columns],
                        use_container_width=True,
                        height=600
                    )
                
                # Show detailed issues for failed records
                if fail_count > 0:
                    st.subheader("❌ Failed Records Details")
                    failed_df = audit_df[audit_df['audit_status'] == 'FAIL']
                    
                    for idx, row in failed_df.iterrows():
                        with st.expander(f"❌ {row.get('FULL_NAME', 'Unknown Guest')} - {row.get('audit_issues', 'No issues listed')}"):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Guest Information:**")
                                st.write(f"Name: {row.get('FULL_NAME', 'N/A')}")
                                st.write(f"Arrival: {row.get('ARRIVAL', 'N/A')}")
                                st.write(f"Departure: {row.get('DEPARTURE', 'N/A')}")
                                st.write(f"Nights: {row.get('NIGHTS', 'N/A')}")
                                st.write(f"Persons: {row.get('PERSONS', 'N/A')}")
                                st.write(f"Room: {row.get('ROOM', 'N/A')}")
                                # Show rate information
                                if pd.notna(row.get('TDF')):
                                    st.write(f"TDF: AED {row.get('TDF', 0):,.2f}")
                                if pd.notna(row.get('NET_TOTAL')):
                                    st.write(f"Net Total: AED {row.get('NET_TOTAL', 0):,.2f}")
                                if pd.notna(row.get('MAIL_TDF_AED')):
                                    st.write(f"Email TDF: {row.get('MAIL_TDF_AED', 'N/A')}")
                                if pd.notna(row.get('MAIL_NET_TOTAL_AED')):
                                    st.write(f"Email Net Total: {row.get('MAIL_NET_TOTAL_AED', 'N/A')}")
                            with col2:
                                st.write("**Issues Found:**")
                                issues = row.get('audit_issues', '').split(';')
                                for issue in issues:
                                    if issue.strip():
                                        st.write(f"• {issue.strip()}")
                
                # Download audit results
                audit_csv = audit_df.to_csv(index=False)
                st.download_button(
                    label="💾 Download Audit Results",
                    data=audit_csv,
                    file_name=f"audit_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            else:
                st.info("👆 Click 'Run Audit Checks' to perform validation on the data.")
        
        else:
            st.info("👆 Go to Email Extraction Results tab to load an Excel file first.")
    
    # Tab 4: Logs & History
    with tab4:
        st.header("📝 Logs & History")
        
        # Recent runs section
        st.subheader("🔄 Recent Runs")
        
        try:
            recent_runs = st.session_state.database.get_recent_runs(limit=20)
            
            if not recent_runs.empty:
                # Summary metrics for recent runs
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Runs", len(recent_runs))
                with col2:
                    completed_runs = len(recent_runs[recent_runs['status'] == 'COMPLETED'])
                    st.metric("Completed", completed_runs)
                with col3:
                    failed_runs = len(recent_runs[recent_runs['status'] == 'FAILED'])
                    st.metric("Failed", failed_runs, delta=f"{failed_runs}" if failed_runs > 0 else None)
                with col4:
                    if st.session_state.current_run_id:
                        st.metric("Current Run", st.session_state.current_run_id[-8:])  # Show last 8 chars
                    else:
                        st.metric("Current Run", "None")
                
                st.markdown("---")
                
                # Runs table with status indicators
                runs_display = recent_runs.copy()
                runs_display['run_timestamp'] = pd.to_datetime(runs_display['run_timestamp']).dt.strftime('%Y-%m-%d %H:%M:%S')
                runs_display['Status'] = runs_display['status'].apply(
                    lambda x: f"🟢 {x}" if x == 'COMPLETED' else f"🔴 {x}" if x == 'FAILED' else f"🟡 {x}"
                )
                
                # Display runs table
                display_columns = ['run_id', 'run_timestamp', 'excel_file_processed', 'Status',
                                 'reservations_loaded_count', 'emails_found_count', 'audit_pass_count', 'audit_fail_count']
                available_display_cols = [col for col in display_columns if col in runs_display.columns]
                
                st.dataframe(
                    runs_display[available_display_cols],
                    use_container_width=True,
                    height=400
                )
                
                # Run details section
                st.subheader("🔍 Run Details")
                
                # Select a run to view details
                selected_run = st.selectbox(
                    "Select a run to view details:",
                    options=recent_runs['run_id'].tolist(),
                    format_func=lambda x: f"{x[-8:]} - {recent_runs[recent_runs['run_id']==x]['run_timestamp'].iloc[0]}"
                )
                
                if selected_run:
                    run_details = recent_runs[recent_runs['run_id'] == selected_run].iloc[0]
                    
                    # Run statistics
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Run Statistics:**")
                        st.write(f"• File: {run_details.get('excel_file_processed', 'N/A')}")
                        st.write(f"• Reservations Loaded: {run_details.get('reservations_loaded_count', 0)}")
                        st.write(f"• Emails Found: {run_details.get('emails_found_count', 0)}")
                        st.write(f"• PDF Extractions: {run_details.get('pdf_extractions_count', 0)}")
                        st.write(f"• Execution Time: {run_details.get('execution_time_seconds', 0):.2f}s")
                    
                    with col2:
                        st.write("**Audit Results:**")
                        st.write(f"• Status: {run_details.get('status', 'Unknown')}")
                        st.write(f"• Passed: {run_details.get('audit_pass_count', 0)}")
                        st.write(f"• Failed: {run_details.get('audit_fail_count', 0)}")
                        if run_details.get('audit_pass_count', 0) + run_details.get('audit_fail_count', 0) > 0:
                            success_rate = (run_details.get('audit_pass_count', 0) / 
                                          (run_details.get('audit_pass_count', 0) + run_details.get('audit_fail_count', 0))) * 100
                            st.write(f"• Success Rate: {success_rate:.1f}%")
                    
                    # Show errors if any
                    errors = st.session_state.database.get_run_errors(selected_run)
                    if errors:
                        st.subheader("❌ Errors & Issues")
                        for idx, error in enumerate(errors, 1):
                            with st.expander(f"Error {idx} - {error.get('timestamp', 'Unknown time')}", expanded=False):
                                st.write(f"**Context:** {error.get('context', 'N/A')}")
                                st.code(error.get('error', 'No error message'), language='text')
                    else:
                        st.success("✅ No errors recorded for this run")
                    
                    # Export options for this run
                    st.subheader("📥 Export Run Data")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        if st.button("Export Raw Data", key=f"export_raw_{selected_run}"):
                            raw_data = st.session_state.database.export_data('reservations_raw', selected_run)
                            if not raw_data.empty:
                                csv = raw_data.to_csv(index=False)
                                st.download_button(
                                    label="💾 Download Raw Data CSV",
                                    data=csv,
                                    file_name=f"raw_data_{selected_run}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv"
                                )
                            else:
                                st.warning("No raw data found for this run")
                    
                    with col2:
                        if st.button("Export Email Data", key=f"export_email_{selected_run}"):
                            email_data = st.session_state.database.export_data('reservations_email', selected_run)
                            if not email_data.empty:
                                csv = email_data.to_csv(index=False)
                                st.download_button(
                                    label="💾 Download Email Data CSV",
                                    data=csv,
                                    file_name=f"email_data_{selected_run}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv"
                                )
                            else:
                                st.warning("No email data found for this run")
                    
                    with col3:
                        if st.button("Export Audit Data", key=f"export_audit_{selected_run}"):
                            audit_data = st.session_state.database.export_data('reservations_audit', selected_run)
                            if not audit_data.empty:
                                csv = audit_data.to_csv(index=False)
                                st.download_button(
                                    label="💾 Download Audit Data CSV",
                                    data=csv,
                                    file_name=f"audit_data_{selected_run}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv"
                                )
                            else:
                                st.warning("No audit data found for this run")
                
            else:
                st.info("📭 No runs found in the database yet. Process some data to see runs here.")
                
            # Database maintenance section
            st.subheader("🧹 Database Maintenance")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🗑️ Clean Old Runs (30+ days)"):
                    try:
                        deleted_count = st.session_state.database.cleanup_old_runs(days_to_keep=30)
                        if deleted_count > 0:
                            st.success(f"✅ Cleaned up {deleted_count} old runs")
                        else:
                            st.info("ℹ️ No old runs to clean up")
                    except Exception as e:
                        st.error(f"❌ Cleanup failed: {e}")
            
            with col2:
                # Database summary stats
                summary_stats = st.session_state.database.get_summary_stats()
                st.write("**Database Summary:**")
                st.write(f"• Total Runs: {summary_stats.get('total_runs', 0)}")
                st.write(f"• Total Audits: {summary_stats.get('total_audits', 0)}")
                
        except Exception as e:
            st.error(f"❌ Error loading logs: {e}")
            st.write("This might be due to database initialization issues. Try processing some data first.")

if __name__ == "__main__":
    main()