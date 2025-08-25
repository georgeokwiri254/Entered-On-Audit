#!/usr/bin/env python3
"""
NER Training Data Extractor
Extracts training data from all MSG files using existing parsers
Converts parser outputs to NER training format
"""

import os
import sys
import json
import csv
import importlib
import traceback
from pathlib import Path
from typing import Dict, List, Any, Tuple
from datetime import datetime
import extract_msg
import re

# Add Rules directories to path for parser imports
sys.path.append('Rules/INNLINKWAY/Agoda')
sys.path.append('Rules/INNLINKWAY/Booking.com')
sys.path.append('Rules/INNLINKWAY/Brand.com')
sys.path.append('Rules/INNLINKWAY/Expedia')

sys.path.extend([
    'Rules/Travel Agency TO/AlKhalidiah',
    'Rules/Travel Agency TO/Almosafer',
    'Rules/Travel Agency TO/Dakkak',
    'Rules/Travel Agency TO/Darina',
    'Rules/Travel Agency TO/Desert Adventures',
    'Rules/Travel Agency TO/Desert Gate',
    'Rules/Travel Agency TO/Dubai Link',
    'Rules/Travel Agency TO/Duri',
    'Rules/Travel Agency TO/Ease My Trip',
    'Rules/Travel Agency TO/Fun&Sun',
    'Rules/Travel Agency TO/Miracle Tourism',
    'Rules/Travel Agency TO/Nirvana',
    'Rules/Travel Agency TO/TBO',
    'Rules/Travel Agency TO/Travco',
    'Rules/Travel Agency TO/Traveltino',
    'Rules/Travel Agency TO/Voyage',
    'Rules/Travel Agency TO/Webbeds'
])

class NERTrainingDataExtractor:
    """Extract and prepare NER training data from MSG files"""
    
    def __init__(self, output_dir: str = "ner_training_data"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Canonical field names for NER
        self.canonical_fields = [
            'MAIL_FIRST_NAME', 'MAIL_FULL_NAME', 'MAIL_ARRIVAL', 'MAIL_DEPARTURE',
            'MAIL_NIGHTS', 'MAIL_PERSONS', 'MAIL_ROOM', 'MAIL_RATE_CODE', 
            'MAIL_C_T_S', 'MAIL_NET_TOTAL', 'MAIL_TOTAL', 'MAIL_TDF', 
            'MAIL_ADR', 'MAIL_AMOUNT'
        ]
        
        # Agency parser mapping
        self.agency_parsers = {
            # INNLINKWAY
            'agoda': {
                'module': 'agoda_parser',
                'extract_func': 'parse_agoda_email',
                'identify_func': None,
                'path': 'Rules/INNLINKWAY/Agoda'
            },
            'booking': {
                'module': 'booking_com_parser',
                'extract_func': 'extract_booking_fields',
                'identify_func': 'is_booking_email',
                'path': 'Rules/INNLINKWAY/Booking.com'
            },
            'brand': {
                'module': 'brand_com_parser',
                'extract_func': 'extract_brand_fields',
                'identify_func': 'is_brand_email',
                'path': 'Rules/INNLINKWAY/Brand.com'
            },
            'expedia': {
                'module': 'expedia_parser',
                'extract_func': 'extract_expedia_fields',
                'identify_func': 'is_expedia_email',
                'path': 'Rules/INNLINKWAY/Expedia'
            },
            
            # Travel Agencies
            'alkhalidiah': {
                'module': 'alkhalidiah_parser',
                'extract_func': 'extract_alkhalidiah_fields',
                'identify_func': 'is_alkhalidiah_email',
                'path': 'Rules/Travel Agency TO/AlKhalidiah'
            },
            'almosafer': {
                'module': 'almosafer_parser',
                'extract_func': 'extract_almosafer_fields',
                'identify_func': 'is_almosafer_email',
                'path': 'Rules/Travel Agency TO/Almosafer'
            },
            'dakkak': {
                'module': 'dakkak_travel_parser',
                'extract_func': 'extract_dakkak_fields',
                'identify_func': 'is_dakkak_email',
                'path': 'Rules/Travel Agency TO/Dakkak'
            },
            'darina': {
                'module': 'darina_parser',
                'extract_func': 'extract_desert_adventures_fields',
                'identify_func': 'is_desert_adventures_file',
                'path': 'Rules/Travel Agency TO/Darina'
            },
            'desert_adventures': {
                'module': 'desert_adventures_parser',
                'extract_func': 'extract_desert_adventures_fields',
                'identify_func': 'is_desert_adventures_file',
                'path': 'Rules/Travel Agency TO/Desert Adventures'
            },
            'desert_gate': {
                'module': 'desert_gate_parser',
                'extract_func': 'extract_desert_gate_fields',
                'identify_func': 'is_desert_gate_email',
                'path': 'Rules/Travel Agency TO/Desert Gate'
            },
            'dubai_link': {
                'module': 'dubai_link_parser',
                'extract_func': 'extract_dubai_link_fields',
                'identify_func': 'is_dubai_link_email',
                'path': 'Rules/Travel Agency TO/Dubai Link'
            },
            'duri': {
                'module': 'duri_parser',
                'extract_func': 'extract_duri_fields',
                'identify_func': 'is_duri_email',
                'path': 'Rules/Travel Agency TO/Duri'
            },
            'ease_my_trip': {
                'module': 'ease_my_trip_parser',
                'extract_func': 'extract_ease_my_trip_fields',
                'identify_func': 'is_ease_my_trip_email',
                'path': 'Rules/Travel Agency TO/Ease My Trip'
            },
            'funsun': {
                'module': 'funsun_parser',
                'extract_func': 'extract_funsun_fields',
                'identify_func': 'is_funsun_file',
                'path': 'Rules/Travel Agency TO/Fun&Sun'
            },
            'miracle_tourism': {
                'module': 'miracle_tourism_parser',
                'extract_func': 'extract_miracle_tourism_fields',
                'identify_func': 'is_miracle_tourism_file',
                'path': 'Rules/Travel Agency TO/Miracle Tourism'
            },
            'nirvana': {
                'module': 'nirvana_parser',
                'extract_func': 'extract_nirvana_fields',
                'identify_func': 'is_nirvana_email',
                'path': 'Rules/Travel Agency TO/Nirvana'
            },
            'tbo': {
                'module': 'tbo_parser',
                'extract_func': 'extract_tbo_fields',
                'identify_func': 'is_tbo_email',
                'path': 'Rules/Travel Agency TO/TBO'
            },
            'travco': {
                'module': 'travco_parser',
                'extract_func': 'extract_travco_fields',
                'identify_func': 'is_travco_email',
                'path': 'Rules/Travel Agency TO/Travco'
            },
            'traveltino': {
                'module': 'traveltino_parser',
                'extract_func': 'extract_traveltino_fields',
                'identify_func': 'is_traveltino_email',
                'path': 'Rules/Travel Agency TO/Traveltino'
            },
            'voyage': {
                'module': 'voyage_parser',
                'extract_func': 'extract_voyage_fields',
                'identify_func': 'is_voyage_email',
                'path': 'Rules/Travel Agency TO/Voyage'
            },
            'webbeds': {
                'module': 'webbeds_parser',
                'extract_func': 'extract_webbeds_fields',
                'identify_func': 'is_webbeds_email',
                'path': 'Rules/Travel Agency TO/Webbeds'
            }
        }
    
    def extract_msg_content(self, msg_path: str) -> Tuple[str, str, str]:
        """Extract content from MSG file"""
        try:
            msg = extract_msg.Message(msg_path)
            
            # Get email content
            body = ""
            if msg.body:
                body = msg.body
            elif msg.htmlBody:
                # Simple HTML to text conversion
                body = re.sub(r'<[^>]+>', '', msg.htmlBody)
            
            subject = msg.subject or ""
            sender = msg.sender or ""
            
            return body, subject, sender
        except Exception as e:
            print(f"Error extracting MSG {msg_path}: {e}")
            return "", "", ""
    
    def identify_agency_from_path(self, msg_path: str) -> str:
        """Identify agency from file path"""
        path_lower = msg_path.lower()
        
        # INNLINKWAY agencies
        if 'agoda' in path_lower:
            return 'agoda'
        elif 'booking.com' in path_lower:
            return 'booking'
        elif 'brand.com' in path_lower:
            return 'brand'
        elif 'expedia' in path_lower:
            return 'expedia'
        
        # Travel agencies
        elif 'alkhalidiah' in path_lower:
            return 'alkhalidiah'
        elif 'almosafer' in path_lower:
            return 'almosafer'
        elif 'dakkak' in path_lower:
            return 'dakkak'
        elif 'darina' in path_lower:
            return 'darina'
        elif 'desert adventures' in path_lower:
            return 'desert_adventures'
        elif 'desert gate' in path_lower:
            return 'desert_gate'
        elif 'dubai link' in path_lower:
            return 'dubai_link'
        elif 'duri' in path_lower:
            return 'duri'
        elif 'ease my trip' in path_lower:
            return 'ease_my_trip'
        elif 'fun&sun' in path_lower:
            return 'funsun'
        elif 'miracle tourism' in path_lower:
            return 'miracle_tourism'
        elif 'nirvana' in path_lower:
            return 'nirvana'
        elif 'tbo' in path_lower:
            return 'tbo'
        elif 'travco' in path_lower:
            return 'travco'
        elif 'traveltino' in path_lower:
            return 'traveltino'
        elif 'voyage' in path_lower:
            return 'voyage'
        elif 'webbeds' in path_lower:
            return 'webbeds'
        
        return 'unknown'
    
    def extract_with_parser(self, agency: str, msg_path: str, body: str, subject: str, sender: str) -> Dict[str, Any]:
        """Extract fields using the appropriate parser"""
        if agency not in self.agency_parsers:
            return {}
        
        parser_info = self.agency_parsers[agency]
        
        try:
            # Import parser module
            module = importlib.import_module(parser_info['module'])
            
            # Get extraction function
            extract_func = getattr(module, parser_info['extract_func'])
            
            # Call extraction function with appropriate parameters
            if agency == 'agoda':
                # Agoda parser has different signature
                fields = extract_func(body, sender)
            elif agency in ['darina', 'desert_adventures', 'funsun', 'miracle_tourism']:
                # File-based parsers
                fields = extract_func(msg_path)
            else:
                # Standard parsers
                fields = extract_func(body, subject)
            
            # Normalize field names to canonical format
            normalized_fields = {}
            for canonical_field in self.canonical_fields:
                # Try direct mapping first
                if canonical_field in fields:
                    normalized_fields[canonical_field] = fields[canonical_field]
                # Try alternative mappings
                elif canonical_field == 'MAIL_FULL_NAME' and 'MAIL_LAST_NAME' in fields:
                    normalized_fields[canonical_field] = fields['MAIL_LAST_NAME']
                elif canonical_field == 'MAIL_C_T_S' and 'C_T_S' in fields:
                    normalized_fields[canonical_field] = fields['C_T_S']
                else:
                    normalized_fields[canonical_field] = 'N/A'
            
            return normalized_fields
            
        except Exception as e:
            print(f"Error extracting with {agency} parser for {msg_path}: {e}")
            traceback.print_exc()
            return {}
    
    def extract_all_training_data(self) -> List[Dict[str, Any]]:
        """Extract training data from all MSG files"""
        training_data = []
        rules_dir = Path("Rules")
        
        # Find all MSG files
        msg_files = list(rules_dir.rglob("*.msg"))
        print(f"Found {len(msg_files)} MSG files")
        
        for msg_path in msg_files:
            print(f"\nProcessing: {msg_path}")
            
            # Extract MSG content
            body, subject, sender = self.extract_msg_content(str(msg_path))
            
            if not body:
                print(f"  ❌ No content extracted")
                continue
            
            # Identify agency
            agency = self.identify_agency_from_path(str(msg_path))
            if agency == 'unknown':
                print(f"  ❌ Unknown agency")
                continue
            
            print(f"  📧 Agency: {agency}")
            print(f"  📄 Content length: {len(body)} chars")
            print(f"  📋 Subject: {subject[:50]}...")
            
            # Extract fields using parser
            fields = self.extract_with_parser(agency, str(msg_path), body, subject, sender)
            
            if not fields:
                print(f"  ❌ No fields extracted")
                continue
            
            # Count extracted fields
            extracted_count = sum(1 for v in fields.values() if v != 'N/A')
            print(f"  ✅ Extracted {extracted_count}/{len(self.canonical_fields)} fields")
            
            # Create training record
            record = {
                'email_id': str(msg_path.name),
                'agency': agency,
                'file_path': str(msg_path),
                'raw_text': body,
                'subject': subject,
                'sender': sender,
                'extraction_timestamp': datetime.now().isoformat(),
                'extracted_fields': fields,
                'field_count': extracted_count
            }
            
            training_data.append(record)
        
        print(f"\n✅ Successfully extracted {len(training_data)} training records")
        return training_data
    
    def save_training_data(self, training_data: List[Dict[str, Any]]):
        """Save training data in multiple formats"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save as JSON
        json_path = self.output_dir / f"training_data_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(training_data, f, indent=2, ensure_ascii=False)
        print(f"📁 Saved JSON: {json_path}")
        
        # Save as CSV summary
        csv_path = self.output_dir / f"training_summary_{timestamp}.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Header
            headers = ['email_id', 'agency', 'field_count', 'subject'] + self.canonical_fields
            writer.writerow(headers)
            
            # Data rows
            for record in training_data:
                row = [
                    record['email_id'],
                    record['agency'],
                    record['field_count'],
                    record['subject'][:100]  # Truncate subject
                ]
                
                # Add field values
                for field in self.canonical_fields:
                    value = record['extracted_fields'].get(field, 'N/A')
                    if isinstance(value, (int, float)):
                        row.append(str(value))
                    else:
                        row.append(str(value)[:50])  # Truncate long values
                
                writer.writerow(row)
        
        print(f"📁 Saved CSV: {csv_path}")
        
        # Generate statistics
        self.generate_statistics(training_data, timestamp)
    
    def generate_statistics(self, training_data: List[Dict[str, Any]], timestamp: str):
        """Generate training data statistics"""
        stats = {
            'total_records': len(training_data),
            'agencies': {},
            'field_coverage': {},
            'extraction_quality': {
                'high_quality': 0,  # 10+ fields extracted
                'medium_quality': 0,  # 5-9 fields extracted
                'low_quality': 0    # <5 fields extracted
            }
        }
        
        # Agency breakdown
        for record in training_data:
            agency = record['agency']
            if agency not in stats['agencies']:
                stats['agencies'][agency] = 0
            stats['agencies'][agency] += 1
            
            # Field coverage
            for field, value in record['extracted_fields'].items():
                if field not in stats['field_coverage']:
                    stats['field_coverage'][field] = {'extracted': 0, 'total': 0}
                
                stats['field_coverage'][field]['total'] += 1
                if value != 'N/A':
                    stats['field_coverage'][field]['extracted'] += 1
            
            # Quality assessment
            field_count = record['field_count']
            if field_count >= 10:
                stats['extraction_quality']['high_quality'] += 1
            elif field_count >= 5:
                stats['extraction_quality']['medium_quality'] += 1
            else:
                stats['extraction_quality']['low_quality'] += 1
        
        # Calculate coverage percentages
        for field_info in stats['field_coverage'].values():
            field_info['coverage_percent'] = (field_info['extracted'] / field_info['total']) * 100
        
        # Save statistics
        stats_path = self.output_dir / f"training_stats_{timestamp}.json"
        with open(stats_path, 'w', encoding='utf-8') as f:
            json.dump(stats, f, indent=2)
        
        print(f"📊 Saved statistics: {stats_path}")
        
        # Print summary
        print("\n" + "="*60)
        print("TRAINING DATA STATISTICS")
        print("="*60)
        print(f"Total Records: {stats['total_records']}")
        print(f"Agencies: {len(stats['agencies'])}")
        
        print("\nQuality Distribution:")
        print(f"  High (10+ fields): {stats['extraction_quality']['high_quality']}")
        print(f"  Medium (5-9 fields): {stats['extraction_quality']['medium_quality']}")
        print(f"  Low (<5 fields): {stats['extraction_quality']['low_quality']}")
        
        print("\nTop Field Coverage:")
        coverage_sorted = sorted(stats['field_coverage'].items(), 
                               key=lambda x: x[1]['coverage_percent'], reverse=True)
        for field, info in coverage_sorted[:8]:
            print(f"  {field}: {info['coverage_percent']:.1f}% ({info['extracted']}/{info['total']})")

def main():
    """Main extraction workflow"""
    print("🚀 Starting NER Training Data Extraction")
    print("="*60)
    
    extractor = NERTrainingDataExtractor()
    
    # Extract training data
    training_data = extractor.extract_all_training_data()
    
    if not training_data:
        print("❌ No training data extracted!")
        return
    
    # Save data
    extractor.save_training_data(training_data)
    
    print("\n🎉 Training data extraction complete!")
    print("Next steps:")
    print("1. Review the generated statistics")
    print("2. Run the BIO converter to prepare NER format")
    print("3. Set up LabelStudio for annotation correction")
    print("4. Train the DistilBERT model on Google Colab")

if __name__ == "__main__":
    main()