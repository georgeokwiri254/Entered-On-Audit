#!/usr/bin/env python3
"""
NER BIO Format Converter
Converts extracted training data to BIO format for NER model training
Handles tokenization and label alignment for transformer models
"""

import json
import re
from pathlib import Path
from typing import List, Dict, Tuple, Any
from datetime import datetime
import random

class NERBIOConverter:
    """Convert training data to BIO format for NER training"""
    
    def __init__(self, input_json: str, output_dir: str = "ner_bio_data"):
        self.input_json = input_json
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Entity labels for BIO tagging
        self.entity_labels = [
            'MAIL_FIRST_NAME', 'MAIL_FULL_NAME', 'MAIL_ARRIVAL', 'MAIL_DEPARTURE',
            'MAIL_NIGHTS', 'MAIL_PERSONS', 'MAIL_ROOM', 'MAIL_RATE_CODE', 
            'MAIL_C_T_S', 'MAIL_NET_TOTAL', 'MAIL_TOTAL', 'MAIL_TDF', 
            'MAIL_ADR', 'MAIL_AMOUNT'
        ]
        
        # Create label set: B-ENTITY, I-ENTITY, O
        self.label_set = ['O']  # Outside
        for entity in self.entity_labels:
            self.label_set.extend([f'B-{entity}', f'I-{entity}'])
        
        # Label to ID mapping
        self.label2id = {label: idx for idx, label in enumerate(self.label_set)}
        self.id2label = {idx: label for label, idx in self.label2id.items()}
        
        print(f"📋 Created {len(self.label_set)} labels: {len(self.entity_labels)} entities + O")
    
    def simple_tokenize(self, text: str) -> List[str]:
        """Simple word-based tokenization"""
        # Split on whitespace and punctuation while preserving them
        tokens = re.findall(r'\w+|[^\w\s]', text)
        return [token for token in tokens if token.strip()]
    
    def find_token_spans(self, tokens: List[str], value: str) -> List[Tuple[int, int]]:
        """Find token spans for a given value in the token list"""
        if not value or value == 'N/A':
            return []
        
        # Clean the value for matching
        value_clean = str(value).strip()
        if not value_clean:
            return []
        
        # Tokenize the value
        value_tokens = self.simple_tokenize(value_clean.lower())
        if not value_tokens:
            return []
        
        spans = []
        tokens_lower = [token.lower() for token in tokens]
        
        # Look for exact sequence matches
        for i in range(len(tokens_lower) - len(value_tokens) + 1):
            if tokens_lower[i:i + len(value_tokens)] == value_tokens:
                spans.append((i, i + len(value_tokens) - 1))
        
        # If no exact match, try fuzzy matching for single tokens
        if not spans and len(value_tokens) == 1:
            target = value_tokens[0]
            for i, token in enumerate(tokens_lower):
                if target in token or token in target:
                    spans.append((i, i))
        
        # For numerical values, try to match just the numbers
        if not spans and re.match(r'^[\d,\.]+$', value_clean):
            number_only = re.sub(r'[^\d\.]', '', value_clean)
            for i, token in enumerate(tokens):
                if number_only in re.sub(r'[^\d\.]', '', token):
                    spans.append((i, i))
        
        return spans
    
    def create_bio_labels(self, tokens: List[str], extracted_fields: Dict[str, Any]) -> List[str]:
        """Create BIO labels for tokens based on extracted fields"""
        labels = ['O'] * len(tokens)
        
        # Track which tokens have been labeled to handle overlaps
        labeled_positions = set()
        
        for entity, value in extracted_fields.items():
            if entity not in self.entity_labels or not value or value == 'N/A':
                continue
            
            # Find token spans for this entity
            spans = self.find_token_spans(tokens, str(value))
            
            for start_idx, end_idx in spans:
                # Check for overlaps with already labeled tokens
                overlap = any(pos in labeled_positions for pos in range(start_idx, end_idx + 1))
                if overlap:
                    continue  # Skip overlapping spans
                
                # Label the span
                if start_idx <= end_idx < len(labels):
                    labels[start_idx] = f'B-{entity}'
                    for j in range(start_idx + 1, end_idx + 1):
                        labels[j] = f'I-{entity}'
                    
                    # Mark positions as labeled
                    labeled_positions.update(range(start_idx, end_idx + 1))
        
        return labels
    
    def convert_record_to_bio(self, record: Dict[str, Any]) -> Dict[str, Any]:
        """Convert a single training record to BIO format"""
        text = record['raw_text']
        extracted_fields = record['extracted_fields']
        
        # Tokenize text
        tokens = self.simple_tokenize(text)
        
        if not tokens:
            return None
        
        # Create BIO labels
        labels = self.create_bio_labels(tokens, extracted_fields)
        
        # Count labeled tokens
        labeled_count = sum(1 for label in labels if label != 'O')
        
        return {
            'email_id': record['email_id'],
            'agency': record['agency'],
            'tokens': tokens,
            'labels': labels,
            'token_count': len(tokens),
            'labeled_token_count': labeled_count,
            'original_fields': extracted_fields,
            'subject': record.get('subject', ''),
            'sender': record.get('sender', '')
        }
    
    def split_data(self, bio_data: List[Dict[str, Any]], 
                   train_ratio: float = 0.8, val_ratio: float = 0.1, test_ratio: float = 0.1) -> Tuple[List, List, List]:
        """Split data into train/validation/test sets"""
        # Shuffle data
        random.seed(42)  # For reproducibility
        shuffled_data = bio_data.copy()
        random.shuffle(shuffled_data)
        
        total = len(shuffled_data)
        train_size = int(total * train_ratio)
        val_size = int(total * val_ratio)
        
        train_data = shuffled_data[:train_size]
        val_data = shuffled_data[train_size:train_size + val_size]
        test_data = shuffled_data[train_size + val_size:]
        
        return train_data, val_data, test_data
    
    def save_conll_format(self, data: List[Dict[str, Any]], filename: str):
        """Save data in CoNLL format"""
        conll_path = self.output_dir / filename
        
        with open(conll_path, 'w', encoding='utf-8') as f:
            for record in data:
                # Write comment with metadata
                f.write(f"# id: {record['email_id']}\n")
                f.write(f"# agency: {record['agency']}\n")
                f.write(f"# tokens: {record['token_count']}, labeled: {record['labeled_token_count']}\n")
                
                # Write tokens and labels
                for token, label in zip(record['tokens'], record['labels']):
                    f.write(f"{token}\t{label}\n")
                
                f.write("\n")  # Empty line between records
        
        print(f"💾 Saved CoNLL format: {conll_path} ({len(data)} records)")
    
    def save_json_format(self, data: List[Dict[str, Any]], filename: str):
        """Save data in JSON format for HuggingFace"""
        json_path = self.output_dir / filename
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        print(f"💾 Saved JSON format: {json_path} ({len(data)} records)")
    
    def save_label_mapping(self):
        """Save label to ID mapping"""
        mapping_path = self.output_dir / "label_mapping.json"
        
        mapping = {
            'label2id': self.label2id,
            'id2label': self.id2label,
            'num_labels': len(self.label_set),
            'entity_labels': self.entity_labels
        }
        
        with open(mapping_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=2)
        
        print(f"🏷️  Saved label mapping: {mapping_path}")
    
    def generate_bio_statistics(self, train_data: List[Dict], val_data: List[Dict], test_data: List[Dict]):
        """Generate statistics for BIO data"""
        stats = {
            'dataset_split': {
                'train': len(train_data),
                'validation': len(val_data),
                'test': len(test_data),
                'total': len(train_data) + len(val_data) + len(test_data)
            },
            'token_statistics': {
                'total_tokens': 0,
                'labeled_tokens': 0,
                'avg_tokens_per_record': 0,
                'avg_labeled_per_record': 0
            },
            'label_distribution': {},
            'agency_distribution': {},
            'entity_coverage': {}
        }
        
        all_data = train_data + val_data + test_data
        
        # Calculate statistics
        total_tokens = 0
        labeled_tokens = 0
        
        for record in all_data:
            total_tokens += record['token_count']
            labeled_tokens += record['labeled_token_count']
            
            # Agency distribution
            agency = record['agency']
            if agency not in stats['agency_distribution']:
                stats['agency_distribution'][agency] = 0
            stats['agency_distribution'][agency] += 1
            
            # Label distribution
            for label in record['labels']:
                if label not in stats['label_distribution']:
                    stats['label_distribution'][label] = 0
                stats['label_distribution'][label] += 1
            
            # Entity coverage
            for entity in self.entity_labels:
                if entity not in stats['entity_coverage']:
                    stats['entity_coverage'][entity] = {'B': 0, 'I': 0, 'total_records': 0}
                
                has_entity = any(label.startswith(f'B-{entity}') or label.startswith(f'I-{entity}') 
                               for label in record['labels'])
                if has_entity:
                    stats['entity_coverage'][entity]['total_records'] += 1
                
                # Count B and I tags
                b_count = sum(1 for label in record['labels'] if label == f'B-{entity}')
                i_count = sum(1 for label in record['labels'] if label == f'I-{entity}')
                
                stats['entity_coverage'][entity]['B'] += b_count
                stats['entity_coverage'][entity]['I'] += i_count
        
        stats['token_statistics']['total_tokens'] = total_tokens
        stats['token_statistics']['labeled_tokens'] = labeled_tokens
        stats['token_statistics']['avg_tokens_per_record'] = total_tokens / len(all_data)
        stats['token_statistics']['avg_labeled_per_record'] = labeled_tokens / len(all_data)
        stats['token_statistics']['labeling_ratio'] = (labeled_tokens / total_tokens) * 100
        
        # Save statistics
        stats_path = self.output_dir / f"bio_statistics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(stats_path, 'w', encoding='utf-8') as f:
            json.dump(stats, f, indent=2)
        
        print(f"📊 Saved BIO statistics: {stats_path}")
        
        # Print summary
        print("\n" + "="*60)
        print("BIO DATA STATISTICS")
        print("="*60)
        
        print(f"Dataset Split:")
        print(f"  Train: {stats['dataset_split']['train']}")
        print(f"  Validation: {stats['dataset_split']['validation']}")
        print(f"  Test: {stats['dataset_split']['test']}")
        print(f"  Total: {stats['dataset_split']['total']}")
        
        print(f"\nToken Statistics:")
        print(f"  Total Tokens: {stats['token_statistics']['total_tokens']:,}")
        print(f"  Labeled Tokens: {stats['token_statistics']['labeled_tokens']:,}")
        print(f"  Labeling Ratio: {stats['token_statistics']['labeling_ratio']:.1f}%")
        print(f"  Avg Tokens/Record: {stats['token_statistics']['avg_tokens_per_record']:.1f}")
        
        print(f"\nTop Entity Coverage:")
        entity_sorted = sorted(stats['entity_coverage'].items(), 
                             key=lambda x: x[1]['total_records'], reverse=True)
        for entity, coverage in entity_sorted[:8]:
            total_mentions = coverage['B'] + coverage['I']
            print(f"  {entity}: {coverage['total_records']} records, {total_mentions} mentions")
    
    def convert_all(self):
        """Main conversion workflow"""
        print(f"🔄 Converting training data to BIO format")
        print(f"📂 Input: {self.input_json}")
        print(f"📁 Output: {self.output_dir}")
        
        # Load training data
        with open(self.input_json, 'r', encoding='utf-8') as f:
            training_data = json.load(f)
        
        print(f"📋 Loaded {len(training_data)} training records")
        
        # Convert to BIO format
        bio_data = []
        failed_conversions = 0
        
        for record in training_data:
            bio_record = self.convert_record_to_bio(record)
            if bio_record:
                bio_data.append(bio_record)
            else:
                failed_conversions += 1
        
        print(f"✅ Converted {len(bio_data)} records to BIO format")
        if failed_conversions:
            print(f"⚠️  Failed to convert {failed_conversions} records")
        
        # Split data
        train_data, val_data, test_data = self.split_data(bio_data)
        print(f"📊 Data split: {len(train_data)} train, {len(val_data)} val, {len(test_data)} test")
        
        # Save in different formats
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # CoNLL format
        self.save_conll_format(train_data, f"train_{timestamp}.conll")
        self.save_conll_format(val_data, f"val_{timestamp}.conll")
        self.save_conll_format(test_data, f"test_{timestamp}.conll")
        
        # JSON format
        self.save_json_format(train_data, f"train_{timestamp}.json")
        self.save_json_format(val_data, f"val_{timestamp}.json")
        self.save_json_format(test_data, f"test_{timestamp}.json")
        
        # Save label mapping
        self.save_label_mapping()
        
        # Generate statistics
        self.generate_bio_statistics(train_data, val_data, test_data)
        
        print(f"\n🎉 BIO conversion complete!")
        return train_data, val_data, test_data

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert training data to BIO format")
    parser.add_argument("input_json", help="Input JSON file from training data extractor")
    parser.add_argument("--output_dir", default="ner_bio_data", help="Output directory")
    
    args = parser.parse_args()
    
    if not Path(args.input_json).exists():
        print(f"❌ Input file not found: {args.input_json}")
        return
    
    converter = NERBIOConverter(args.input_json, args.output_dir)
    converter.convert_all()
    
    print("\nNext steps:")
    print("1. Review the generated BIO statistics")
    print("2. Upload data to LabelStudio for annotation correction")
    print("3. Use the JSON files for training with HuggingFace transformers")

if __name__ == "__main__":
    main()