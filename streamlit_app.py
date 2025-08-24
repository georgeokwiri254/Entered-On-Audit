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
    """Extract text from PDF bytes"""
    try:
        pdf_file = io.BytesIO(pdf_bytes)
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except Exception as e:
        logger.error(f"PDF extraction failed: {e}")
        return ""

def extract_reservation_fields(text, sender_email=""):
    """Extract reservation fields using regex patterns"""
    
    # Different patterns for different email sources
    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
        # Patterns for noreply-reservations emails based on actual structure
        patterns = {
            'GUEST_NAME_FULL': r"Guest Name:\s*(.+?)(?:\n|Address:)",
            'ARRIVAL': r"Arrive:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'DEPARTURE': r"Depart:\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'NIGHTS': r"Total Nights\s*(\d+)\s*night",
            'PERSONS': r"Adult/Children:\s*(\d+)/\d+",
            'ROOM': r"Room with One King Bed\n(?:.*\n)?Room Code:\s*[A-Z0-9]+\s*(?:.*\n)?",
            'ROOM_TYPE': r"(Superior Room|Deluxe Room|Standard Room|Executive Room)",
            'RATE_CODE': r"Rate Code:\s*([A-Z0-9]+)",
            'RATE_NAME': r"Rate Name:\s*(.+?)(?:\n|Rate Code:)",
            'COMPANY': r"Travel Agent\s*(?:.*\n)*Name:\s*(.+?)(?:\n|$)",
            'NET_TOTAL': r"Total charges:\s*AED\s*([0-9,]+\.?[0-9]*)",
            'CONFIRMATION': r"Confirman:\s*([A-Z0-9]+)",
        }
        
        # Additional patterns to extract from subject line
        subject_patterns = {
            'ARRIVAL_SUBJECT': r"Arrival Date[:]*\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'CONFIRMATION_SUBJECT': r"confirmation number[:]*\s*([A-Z0-9]+)",
        }
        
        # Merge subject patterns with main patterns
        patterns.update(subject_patterns)
    else:
        # Original patterns for other emails (PDFs, etc.)
        patterns = {
            'FULL_NAME': r"(?:Name|Guest Name)[:\s]+(.+?)(?:\n|$)",
            'FIRST_NAME': r"(?:First Name)[:\s]+(.+?)(?:\n|$)",
            'ARRIVAL': r"(?:Arrival|Check-in)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'DEPARTURE': r"(?:Departure|Check-out)[:\s]+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",
            'NIGHTS': r"(?:Nights|Night)[:\s]+(\d+)",
            'PERSONS': r"(?:Persons|Guest|Adults?)[:\s]+(\d+)",
            'ROOM': r"(?:Room|Room Type)[:\s]+(.+?)(?:\n|$)",
            'RATE_CODE': r"(?:Rate Code|Rate)[:\s]+(.+?)(?:\n|$)",
            'COMPANY': r"(?:Company|Agency)[:\s]+(.+?)(?:\n|$)",
            'NET_TOTAL': r"(?:Total|Net Total|Amount|Net Amount)[:\s]+(?:AED\s*)?([\\d,]+\.?\\d*)",
        }
    
    extracted = {}
    
    # Extract all fields first
    for field, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
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
        
        # Map room types to codes
        room_type = extracted.get('ROOM_TYPE', 'N/A')
        if room_type != 'N/A':
            if 'Superior Room' in room_type:
                extracted['ROOM'] = 'SK'  # Map Superior Room to SK as requested
            else:
                extracted['ROOM'] = room_type
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
    
    # Convert dates to dd/mm/yyyy format
    for date_field in ['ARRIVAL', 'DEPARTURE', 'ARRIVAL_SUBJECT']:
        if date_field in extracted and extracted[date_field] != 'N/A':
            try:
                # Always use dayfirst=True to ensure dd/mm/yyyy interpretation
                parsed_date = pd.to_datetime(extracted[date_field], dayfirst=True)
                extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
            except:
                # If parsing fails, try different formats
                try:
                    # Try mm/dd/yyyy format as fallback
                    parsed_date = pd.to_datetime(extracted[date_field], dayfirst=False)
                    extracted[date_field] = parsed_date.strftime('%d/%m/%Y')
                except:
                    pass  # Keep original value if all parsing fails
    
    # Use arrival from subject if main arrival not found
    if extracted.get('ARRIVAL', 'N/A') == 'N/A' and extracted.get('ARRIVAL_SUBJECT', 'N/A') != 'N/A':
        extracted['ARRIVAL'] = extracted['ARRIVAL_SUBJECT']
    
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
    except:
        extracted['TDF'] = "N/A"
    
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
    
    # Set AMOUNT = NET_TOTAL for consistency
    try:
        net_total = extracted.get('NET_TOTAL', 'N/A')
        if net_total != 'N/A':
            amount_num = float(str(net_total).replace(',', ''))
            extracted['AMOUNT'] = net_total
            extracted['AMOUNT_AED'] = f"AED {amount_num:,.2f}"
            # Set TOTAL as separate field (could be different from NET_TOTAL)
            extracted['TOTAL'] = net_total  # For now, same as NET_TOTAL
        else:
            extracted['AMOUNT'] = "N/A"
            extracted['TOTAL'] = "N/A"
    except:
        extracted['AMOUNT'] = "N/A"
        extracted['TOTAL'] = "N/A"
    
    # Map COMPANY to C_T_S (Company name)
    if extracted.get('COMPANY', 'N/A') != 'N/A':
        extracted['C_T_S'] = extracted['COMPANY']
    else:
        extracted['C_T_S'] = "N/A"
    
    return extracted


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
                                # Save attachment temporarily
                                temp_path = f"temp_{filename.replace(' ', '_').replace('/', '_')}"
                                attachment.SaveAsFile(temp_path)
                                
                                with open(temp_path, 'rb') as f:
                                    pdf_data = f.read()
                                    text = extract_pdf_text(pdf_data)
                                    
                                    if text:
                                        extracted_fields = extract_reservation_fields(text, sender_email)
                                        
                                        # Format currency fields
                                        for field in ['NET_TOTAL', 'TDF']:
                                            if extracted_fields.get(field) != 'N/A':
                                                try:
                                                    amount = float(extracted_fields[field].replace(',', ''))
                                                    extracted_fields[f'{field}_AED'] = f"AED {amount:,.2f}"
                                                except:
                                                    pass
                                        
                                        email_info['extracted_data'] = extracted_fields
                                        email_info['attachments'].append({
                                            'filename': filename,
                                            'size': len(pdf_data),
                                            'text_extracted': bool(text)
                                        })
                                
                                # Clean up temp file
                                if os.path.exists(temp_path):
                                    os.remove(temp_path)
                                    
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
                    
                    # Also extract from body text for noreply-reservations emails
                    sender_email = getattr(email, 'sender', '')
                    if "noreply-reservations@millenniumhotels.com" in sender_email.lower():
                        # Get subject and body content
                        email_text = f"{email.get('subject', '')}\n{getattr(email, 'body', '')}"
                        additional_fields = extract_reservation_fields(email_text, sender_email)
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
                 'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
    
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
                         'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
            
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
        st.write("• Enhanced current mailbox search (all folders)")
        st.write("• Updated Email Extraction Results format to match Entered On sheet")
        st.write("• Enhanced search for names like 'Avital' and 'Shi Guang'")
        st.write("• Support for noreply-reservations@millenniumhotels.com emails")
        st.write("• Added comprehensive rate extraction (ADR, TDF calculation)")
        st.write("• Currency set to AED only")
        st.write("• Email vs data field matching audit")
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
            
            # Email Extraction Results Table - NO DROPDOWNS, JUST TABLE
            st.subheader("📄 Email Extraction Results")
            
            # Create table with extracted email data - NO Guest_Name column, separate NET_TOTAL and TOTAL
            table_data = []
            # Entered On sheet columns A-P (removing Guest_Name, separating NET_TOTAL and TOTAL, C_T_S as Company name)
            specified_fields = ['FULL_NAME', 'FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 
                              'NIGHTS', 'PERSONS', 'ROOM', 'RATE_CODE', 'C_T_S',
                              'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT', 'SEASON']
            
            for result in filtered_results:
                reservation = result['reservation_data']
                guest_name = reservation.get('FULL_NAME', 'N/A')
                
                # Get all extracted email data for this guest
                email_data = {}
                for email in result.get('matching_emails', []):
                    if email.get('extracted_data'):
                        email_data.update(email['extracted_data'])
                
                # Create row with all specified fields - NO Guest_Name column
                row_data = {
                    'Email_Status': result['status'],
                    'Emails_Found': result['email_count']
                }
                
                # Add all specified fields - show N/A if not found
                for field in specified_fields:
                    value = email_data.get(field, 'N/A')
                    # Format currency fields
                    if field in ['TDF', 'NET_TOTAL', 'TOTAL', 'AMOUNT'] and value != 'N/A':
                        try:
                            amount = float(str(value).replace(',', ''))
                            value = f"AED {amount:,.2f}"
                        except:
                            value = 'N/A'
                    row_data[field] = value
                
                # Add corresponding Mail_ columns for extracted email data
                mail_fields = ['FIRST_NAME', 'ARRIVAL', 'DEPARTURE', 'NIGHTS', 'PERSONS', 'ROOM', 'RATE_CODE', 
                             'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT', 'SEASON']
                for field in mail_fields:
                    mail_value = email_data.get(field, 'N/A')
                    # Format currency fields
                    if field in ['TDF', 'NET_TOTAL', 'TOTAL', 'AMOUNT'] and mail_value != 'N/A':
                        try:
                            amount = float(str(mail_value).replace(',', ''))
                            mail_value = f"AED {amount:,.2f}"
                        except:
                            mail_value = 'N/A'
                    row_data[f'Mail_{field}'] = mail_value
                
                table_data.append(row_data)
            
            if table_data:
                results_df = pd.DataFrame(table_data)
                st.dataframe(results_df, use_container_width=True, height=500)
            else:
                st.info("No results to display.")
            
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
            st.info("👆 Load an Excel file using the options above, then click 'Search Emails for Each Reservation' in the sidebar.")
    
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
            st.info("👆 Upload an Excel file in the sidebar to see the converted data.")
    
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
                                   'RATE_CODE', 'C_T_S', 'NET_TOTAL', 'TOTAL', 'TDF', 'ADR', 'AMOUNT']
                
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
                            if original_val != 'N/A' and mail_val != 'N/A' and original_val != mail_val:
                                # Apply red text to Mail_ column if values don't match
                                try:
                                    mail_col_idx = row.index.get_loc(mail_col)
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
            st.info("👆 Upload an Excel file in the sidebar first to run audit checks.")
    
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