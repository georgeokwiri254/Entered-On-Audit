#!/usr/bin/env python3
"""
NER Training Data Validator
Validates training data quality and provides statistics for NER training
"""

import json
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from typing import Dict, List, Tuple, Any
from collections import Counter, defaultdict
import pandas as pd
from datetime import datetime
import re

class NERTrainingValidator:
    """Validate and analyze NER training data quality"""
    
    def __init__(self, training_data_path: str, bio_data_path: str = None):
        self.training_data_path = training_data_path
        self.bio_data_path = bio_data_path
        self.output_dir = Path("ner_validation_results")
        self.output_dir.mkdir(exist_ok=True)
        
        # Load data
        self.training_data = self.load_json(training_data_path)
        self.bio_data = self.load_json(bio_data_path) if bio_data_path else None
        
        # Entity labels
        self.entity_labels = [
            'MAIL_FIRST_NAME', 'MAIL_FULL_NAME', 'MAIL_ARRIVAL', 'MAIL_DEPARTURE',
            'MAIL_NIGHTS', 'MAIL_PERSONS', 'MAIL_ROOM', 'MAIL_RATE_CODE', 
            'MAIL_C_T_S', 'MAIL_NET_TOTAL', 'MAIL_TOTAL', 'MAIL_TDF', 
            'MAIL_ADR', 'MAIL_AMOUNT'
        ]
        
        print(f"📊 Loaded {len(self.training_data)} training records")
        if self.bio_data:
            print(f"📊 Loaded {len(self.bio_data)} BIO records")
    
    def load_json(self, file_path: str) -> List[Dict]:
        """Load JSON data safely"""
        if not file_path or not Path(file_path).exists():
            return []
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ Error loading {file_path}: {e}")
            return []
    
    def validate_extraction_quality(self) -> Dict[str, Any]:
        """Validate extraction quality from parsers"""
        print("🔍 Validating extraction quality...")
        
        validation_results = {
            'total_records': len(self.training_data),
            'quality_tiers': {
                'excellent': 0,    # 12-14 fields extracted
                'good': 0,         # 8-11 fields extracted  
                'fair': 0,         # 4-7 fields extracted
                'poor': 0          # <4 fields extracted
            },
            'field_coverage': {},
            'agency_performance': {},
            'common_issues': [],
            'data_completeness': {}
        }
        
        # Analyze each record
        for record in self.training_data:
            extracted_fields = record.get('extracted_fields', {})
            agency = record.get('agency', 'unknown')
            field_count = record.get('field_count', 0)
            
            # Quality tiers
            if field_count >= 12:
                validation_results['quality_tiers']['excellent'] += 1
            elif field_count >= 8:
                validation_results['quality_tiers']['good'] += 1
            elif field_count >= 4:
                validation_results['quality_tiers']['fair'] += 1
            else:
                validation_results['quality_tiers']['poor'] += 1
            
            # Agency performance tracking
            if agency not in validation_results['agency_performance']:
                validation_results['agency_performance'][agency] = {
                    'total_records': 0,
                    'total_fields_extracted': 0,
                    'avg_fields_per_record': 0,
                    'field_counts': []
                }
            
            validation_results['agency_performance'][agency]['total_records'] += 1
            validation_results['agency_performance'][agency]['total_fields_extracted'] += field_count
            validation_results['agency_performance'][agency]['field_counts'].append(field_count)
            
            # Field coverage analysis
            for field in self.entity_labels:
                if field not in validation_results['field_coverage']:
                    validation_results['field_coverage'][field] = {
                        'extracted_count': 0,
                        'total_records': 0,
                        'coverage_percent': 0,
                        'sample_values': []
                    }
                
                validation_results['field_coverage'][field]['total_records'] += 1
                
                if field in extracted_fields and extracted_fields[field] != 'N/A':
                    validation_results['field_coverage'][field]['extracted_count'] += 1
                    
                    # Collect sample values (first 5 unique)
                    value = str(extracted_fields[field])
                    samples = validation_results['field_coverage'][field]['sample_values']
                    if len(samples) < 5 and value not in samples:
                        samples.append(value)
        
        # Calculate averages and percentages
        for agency, stats in validation_results['agency_performance'].items():
            if stats['total_records'] > 0:
                stats['avg_fields_per_record'] = stats['total_fields_extracted'] / stats['total_records']
        
        for field, stats in validation_results['field_coverage'].items():
            if stats['total_records'] > 0:
                stats['coverage_percent'] = (stats['extracted_count'] / stats['total_records']) * 100
        
        # Identify common issues
        validation_results['common_issues'] = self.identify_common_issues()
        
        return validation_results
    
    def identify_common_issues(self) -> List[Dict[str, Any]]:
        """Identify common data quality issues"""
        issues = []
        
        date_pattern = re.compile(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{4}')
        amount_pattern = re.compile(r'[\d,]+\.?\d*')
        
        for record in self.training_data:
            extracted_fields = record.get('extracted_fields', {})
            email_id = record.get('email_id', 'unknown')
            
            # Check for date format consistency
            arrival = extracted_fields.get('MAIL_ARRIVAL', '')
            departure = extracted_fields.get('MAIL_DEPARTURE', '')
            
            if arrival != 'N/A' and not date_pattern.match(str(arrival)):
                issues.append({
                    'type': 'invalid_date_format',
                    'field': 'MAIL_ARRIVAL',
                    'value': arrival,
                    'email_id': email_id,
                    'severity': 'medium'
                })
            
            if departure != 'N/A' and not date_pattern.match(str(departure)):
                issues.append({
                    'type': 'invalid_date_format', 
                    'field': 'MAIL_DEPARTURE',
                    'value': departure,
                    'email_id': email_id,
                    'severity': 'medium'
                })
            
            # Check for missing critical fields
            critical_fields = ['MAIL_FIRST_NAME', 'MAIL_ARRIVAL', 'MAIL_DEPARTURE', 'MAIL_C_T_S']
            missing_critical = [field for field in critical_fields 
                               if extracted_fields.get(field, 'N/A') == 'N/A']
            
            if len(missing_critical) >= 3:  # More than half critical fields missing
                issues.append({
                    'type': 'missing_critical_fields',
                    'fields': missing_critical,
                    'email_id': email_id,
                    'severity': 'high'
                })
            
            # Check for suspicious amount values
            for amount_field in ['MAIL_NET_TOTAL', 'MAIL_TOTAL', 'MAIL_AMOUNT']:
                amount_value = extracted_fields.get(amount_field, 'N/A')
                if amount_value != 'N/A':
                    try:
                        float_val = float(str(amount_value).replace(',', ''))
                        if float_val <= 0 or float_val > 50000:  # Suspicious range
                            issues.append({
                                'type': 'suspicious_amount',
                                'field': amount_field,
                                'value': amount_value,
                                'email_id': email_id,
                                'severity': 'low'
                            })
                    except (ValueError, TypeError):
                        issues.append({
                            'type': 'invalid_amount_format',
                            'field': amount_field,
                            'value': amount_value,
                            'email_id': email_id,
                            'severity': 'medium'
                        })
        
        return issues
    
    def validate_bio_format(self) -> Dict[str, Any]:
        """Validate BIO format data"""
        if not self.bio_data:
            return {'error': 'No BIO data provided'}
        
        print("🔍 Validating BIO format...")
        
        validation_results = {
            'total_records': len(self.bio_data),
            'label_statistics': {},
            'sequence_statistics': {
                'avg_tokens_per_record': 0,
                'avg_labeled_tokens_per_record': 0,
                'total_tokens': 0,
                'total_labeled_tokens': 0,
                'labeling_ratio': 0
            },
            'bio_format_issues': [],
            'entity_statistics': {}
        }
        
        total_tokens = 0
        total_labeled_tokens = 0
        label_counter = Counter()
        
        # Analyze each BIO record
        for i, record in enumerate(self.bio_data):
            tokens = record.get('tokens', [])
            labels = record.get('labels', [])
            
            if len(tokens) != len(labels):
                validation_results['bio_format_issues'].append({
                    'type': 'token_label_mismatch',
                    'record_index': i,
                    'tokens_count': len(tokens),
                    'labels_count': len(labels),
                    'severity': 'high'
                })
                continue
            
            total_tokens += len(tokens)
            labeled_count = sum(1 for label in labels if label != 'O')
            total_labeled_tokens += labeled_count
            
            # Count labels
            for label in labels:
                label_counter[label] += 1
            
            # Validate BIO format consistency
            self.validate_bio_sequence(labels, i, validation_results['bio_format_issues'])
        
        # Calculate statistics
        if len(self.bio_data) > 0:
            validation_results['sequence_statistics']['avg_tokens_per_record'] = total_tokens / len(self.bio_data)
            validation_results['sequence_statistics']['avg_labeled_tokens_per_record'] = total_labeled_tokens / len(self.bio_data)
        
        validation_results['sequence_statistics']['total_tokens'] = total_tokens
        validation_results['sequence_statistics']['total_labeled_tokens'] = total_labeled_tokens
        
        if total_tokens > 0:
            validation_results['sequence_statistics']['labeling_ratio'] = (total_labeled_tokens / total_tokens) * 100
        
        validation_results['label_statistics'] = dict(label_counter)
        
        # Entity statistics
        for entity in self.entity_labels:
            b_count = label_counter.get(f'B-{entity}', 0)
            i_count = label_counter.get(f'I-{entity}', 0)
            validation_results['entity_statistics'][entity] = {
                'entity_mentions': b_count,
                'total_tokens': b_count + i_count,
                'avg_tokens_per_mention': (b_count + i_count) / b_count if b_count > 0 else 0
            }
        
        return validation_results
    
    def validate_bio_sequence(self, labels: List[str], record_index: int, issues: List[Dict]):
        """Validate BIO sequence consistency"""
        for i, label in enumerate(labels):
            if label.startswith('I-'):
                entity = label[2:]  # Remove I- prefix
                
                # Check if I- tag follows B- or I- of same entity
                if i == 0 or (not labels[i-1].endswith(f'-{entity}')):
                    issues.append({
                        'type': 'invalid_bio_sequence',
                        'record_index': record_index,
                        'token_index': i,
                        'label': label,
                        'previous_label': labels[i-1] if i > 0 else 'START',
                        'severity': 'medium'
                    })
    
    def generate_visualizations(self, validation_results: Dict[str, Any]):
        """Generate validation visualizations"""
        print("📊 Generating visualizations...")
        
        # Set style
        plt.style.use('seaborn-v0_8')
        fig, axes = plt.subplots(2, 2, figsize=(15, 12))
        fig.suptitle('NER Training Data Validation Report', fontsize=16, fontweight='bold')
        
        # 1. Quality Distribution Pie Chart
        quality_data = validation_results['quality_tiers']
        axes[0, 0].pie(quality_data.values(), labels=quality_data.keys(), autopct='%1.1f%%', startangle=90)
        axes[0, 0].set_title('Data Quality Distribution')
        
        # 2. Field Coverage Bar Chart
        field_coverage = validation_results['field_coverage']
        fields = list(field_coverage.keys())
        coverages = [field_coverage[field]['coverage_percent'] for field in fields]
        
        bars = axes[0, 1].bar(range(len(fields)), coverages, color='skyblue')
        axes[0, 1].set_title('Field Coverage Percentage')
        axes[0, 1].set_xlabel('Entity Fields')
        axes[0, 1].set_ylabel('Coverage %')
        axes[0, 1].set_xticks(range(len(fields)))
        axes[0, 1].set_xticklabels(fields, rotation=45, ha='right')
        
        # Add value labels on bars
        for bar, coverage in zip(bars, coverages):
            height = bar.get_height()
            axes[0, 1].text(bar.get_x() + bar.get_width()/2., height + 1,
                           f'{coverage:.1f}%', ha='center', va='bottom', fontsize=8)
        
        # 3. Agency Performance
        agency_performance = validation_results['agency_performance']
        agencies = list(agency_performance.keys())[:10]  # Top 10 agencies
        avg_fields = [agency_performance[agency]['avg_fields_per_record'] for agency in agencies]
        
        bars = axes[1, 0].barh(agencies, avg_fields, color='lightgreen')
        axes[1, 0].set_title('Agency Performance (Avg Fields Extracted)')
        axes[1, 0].set_xlabel('Average Fields per Email')
        
        # Add value labels
        for bar, avg_field in zip(bars, avg_fields):
            width = bar.get_width()
            axes[1, 0].text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                           f'{avg_field:.1f}', ha='left', va='center')
        
        # 4. Issues Severity Distribution
        issues = validation_results.get('common_issues', [])
        severity_counts = Counter(issue['severity'] for issue in issues)
        
        if severity_counts:
            axes[1, 1].bar(severity_counts.keys(), severity_counts.values(), 
                          color=['red', 'orange', 'yellow'])
            axes[1, 1].set_title('Data Quality Issues by Severity')
            axes[1, 1].set_ylabel('Issue Count')
            
            # Add value labels
            for i, (severity, count) in enumerate(severity_counts.items()):
                axes[1, 1].text(i, count + 0.5, str(count), ha='center', va='bottom')
        else:
            axes[1, 1].text(0.5, 0.5, 'No issues detected', ha='center', va='center',
                           transform=axes[1, 1].transAxes, fontsize=12)
            axes[1, 1].set_title('Data Quality Issues by Severity')
        
        plt.tight_layout()
        
        # Save plot
        plot_path = self.output_dir / f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"📊 Saved validation plot: {plot_path}")
        return str(plot_path)
    
    def generate_detailed_report(self, validation_results: Dict[str, Any], 
                               bio_results: Dict[str, Any] = None) -> str:
        """Generate detailed validation report"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = self.output_dir / f"validation_report_{timestamp}.md"
        
        report = f"""# NER Training Data Validation Report
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 📊 Overview
- **Total Records**: {validation_results['total_records']:,}
- **Data Source**: {self.training_data_path}
- **Entity Labels**: {len(self.entity_labels)}

## 🎯 Quality Distribution
"""
        
        # Quality tiers
        quality_data = validation_results['quality_tiers']
        total_records = validation_results['total_records']
        
        for tier, count in quality_data.items():
            percentage = (count / total_records) * 100 if total_records > 0 else 0
            report += f"- **{tier.title()}**: {count:,} records ({percentage:.1f}%)\n"
        
        report += "\n## 📋 Field Coverage Analysis\n\n"
        report += "| Entity | Coverage % | Extracted | Total | Sample Values |\n"
        report += "|--------|------------|-----------|-------|--------------|\n"
        
        # Sort by coverage percentage
        field_coverage = validation_results['field_coverage']
        sorted_fields = sorted(field_coverage.items(), 
                             key=lambda x: x[1]['coverage_percent'], reverse=True)
        
        for field, stats in sorted_fields:
            samples = ', '.join(stats['sample_values'][:3])  # First 3 samples
            if len(stats['sample_values']) > 3:
                samples += '...'
            
            report += f"| {field} | {stats['coverage_percent']:.1f}% | {stats['extracted_count']:,} | {stats['total_records']:,} | {samples} |\n"
        
        report += "\n## 🏢 Agency Performance\n\n"
        report += "| Agency | Records | Avg Fields | Performance |\n"
        report += "|--------|---------|------------|-------------|\n"
        
        # Sort agencies by performance
        agency_performance = validation_results['agency_performance']
        sorted_agencies = sorted(agency_performance.items(), 
                               key=lambda x: x[1]['avg_fields_per_record'], reverse=True)
        
        for agency, stats in sorted_agencies:
            performance = "🟢 Excellent" if stats['avg_fields_per_record'] >= 10 else \
                         "🟡 Good" if stats['avg_fields_per_record'] >= 6 else \
                         "🔴 Needs Improvement"
            
            report += f"| {agency} | {stats['total_records']:,} | {stats['avg_fields_per_record']:.1f} | {performance} |\n"
        
        # Issues section
        issues = validation_results.get('common_issues', [])
        if issues:
            report += "\n## ⚠️ Data Quality Issues\n\n"
            
            issue_types = defaultdict(list)
            for issue in issues:
                issue_types[issue['type']].append(issue)
            
            for issue_type, issue_list in issue_types.items():
                report += f"### {issue_type.replace('_', ' ').title()}\n"
                report += f"**Count**: {len(issue_list)} issues\n\n"
                
                # Show first few examples
                for issue in issue_list[:5]:
                    report += f"- **{issue.get('email_id', 'unknown')}**: "
                    if 'field' in issue:
                        report += f"{issue['field']} = '{issue.get('value', 'N/A')}'"
                    elif 'fields' in issue:
                        report += f"Missing: {', '.join(issue['fields'])}"
                    report += f" (Severity: {issue['severity']})\n"
                
                if len(issue_list) > 5:
                    report += f"- ... and {len(issue_list) - 5} more\n"
                
                report += "\n"
        
        # BIO format validation if available
        if bio_results and 'error' not in bio_results:
            report += "\n## 🏷️ BIO Format Validation\n\n"
            
            seq_stats = bio_results['sequence_statistics']
            report += f"- **Total Tokens**: {seq_stats['total_tokens']:,}\n"
            report += f"- **Labeled Tokens**: {seq_stats['total_labeled_tokens']:,}\n"
            report += f"- **Labeling Ratio**: {seq_stats['labeling_ratio']:.1f}%\n"
            report += f"- **Avg Tokens/Record**: {seq_stats['avg_tokens_per_record']:.1f}\n"
            report += f"- **Avg Labeled/Record**: {seq_stats['avg_labeled_tokens_per_record']:.1f}\n"
            
            # Entity statistics
            report += "\n### Entity Statistics\n\n"
            report += "| Entity | Mentions | Total Tokens | Avg Tokens/Mention |\n"
            report += "|--------|----------|--------------|--------------------|\n"
            
            entity_stats = bio_results['entity_statistics']
            for entity, stats in entity_stats.items():
                if stats['entity_mentions'] > 0:
                    report += f"| {entity} | {stats['entity_mentions']:,} | {stats['total_tokens']:,} | {stats['avg_tokens_per_mention']:.1f} |\n"
            
            # BIO format issues
            bio_issues = bio_results.get('bio_format_issues', [])
            if bio_issues:
                report += f"\n### BIO Format Issues ({len(bio_issues)} found)\n\n"
                issue_types = defaultdict(list)
                for issue in bio_issues:
                    issue_types[issue['type']].append(issue)
                
                for issue_type, issue_list in issue_types.items():
                    report += f"- **{issue_type.replace('_', ' ').title()}**: {len(issue_list)} issues\n"
        
        report += "\n## 🎯 Recommendations\n\n"
        
        # Generate recommendations based on analysis
        recommendations = self.generate_recommendations(validation_results, bio_results)
        for rec in recommendations:
            report += f"- {rec}\n"
        
        report += f"\n## 📁 Files Analyzed\n\n"
        report += f"- **Training Data**: {self.training_data_path}\n"
        if self.bio_data_path:
            report += f"- **BIO Data**: {self.bio_data_path}\n"
        
        # Save report
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        print(f"📄 Generated detailed report: {report_path}")
        return str(report_path)
    
    def generate_recommendations(self, validation_results: Dict, bio_results: Dict = None) -> List[str]:
        """Generate actionable recommendations"""
        recommendations = []
        
        # Quality-based recommendations
        quality = validation_results['quality_tiers']
        total = validation_results['total_records']
        
        if quality['poor'] / total > 0.2:  # More than 20% poor quality
            recommendations.append("🔧 **Improve Parser Quality**: >20% of records have poor extraction. Review and enhance parser regex patterns.")
        
        if quality['excellent'] / total < 0.3:  # Less than 30% excellent
            recommendations.append("📈 **Enhance Field Coverage**: Focus on improving extraction for underperforming entities.")
        
        # Field coverage recommendations
        field_coverage = validation_results['field_coverage']
        low_coverage_fields = [field for field, stats in field_coverage.items() 
                              if stats['coverage_percent'] < 50]
        
        if low_coverage_fields:
            recommendations.append(f"🎯 **Focus on Low Coverage Fields**: {', '.join(low_coverage_fields[:5])} need attention.")
        
        # Agency-specific recommendations
        agency_performance = validation_results['agency_performance']
        poor_agencies = [agency for agency, stats in agency_performance.items() 
                        if stats['avg_fields_per_record'] < 4]
        
        if poor_agencies:
            recommendations.append(f"🏢 **Improve Agency Parsers**: {', '.join(poor_agencies[:3])} parsers need enhancement.")
        
        # Data quality issues
        issues = validation_results.get('common_issues', [])
        high_severity = [i for i in issues if i['severity'] == 'high']
        
        if high_severity:
            recommendations.append("🚨 **Address High Severity Issues**: Fix critical field extraction problems before training.")
        
        # BIO format recommendations
        if bio_results and 'error' not in bio_results:
            labeling_ratio = bio_results['sequence_statistics']['labeling_ratio']
            
            if labeling_ratio < 5:
                recommendations.append("🏷️ **Increase Label Density**: Current labeling ratio is low. Consider adding more entity annotations.")
            
            bio_issues = bio_results.get('bio_format_issues', [])
            if bio_issues:
                recommendations.append("🔍 **Fix BIO Format Issues**: Resolve sequence consistency problems before training.")
        
        # General recommendations
        if total < 50:
            recommendations.append("📊 **Increase Data Size**: Consider collecting more training samples for better model performance.")
        
        recommendations.append("✅ **Manual Review**: Review samples with LabelStudio to correct extraction errors.")
        recommendations.append("🧪 **Test Split**: Reserve 15-20% of data for testing to evaluate model performance.")
        
        return recommendations
    
    def run_full_validation(self) -> Dict[str, Any]:
        """Run complete validation workflow"""
        print("🚀 Starting full NER training validation...")
        print("=" * 60)
        
        # Validate extraction quality
        validation_results = self.validate_extraction_quality()
        
        # Validate BIO format if available
        bio_results = self.validate_bio_format() if self.bio_data else None
        
        # Generate visualizations
        plot_path = self.generate_visualizations(validation_results)
        
        # Generate detailed report
        report_path = self.generate_detailed_report(validation_results, bio_results)
        
        # Summary
        results_summary = {
            'validation_results': validation_results,
            'bio_results': bio_results,
            'plot_path': plot_path,
            'report_path': report_path,
            'recommendations': self.generate_recommendations(validation_results, bio_results)
        }
        
        # Save summary
        summary_path = self.output_dir / f"validation_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(results_summary, f, indent=2, default=str)
        
        print("\n🎉 Validation Complete!")
        print("=" * 60)
        print(f"📊 Report: {report_path}")
        print(f"📈 Plot: {plot_path}")
        print(f"📋 Summary: {summary_path}")
        
        return results_summary

def main():
    """Main validation workflow"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Validate NER training data")
    parser.add_argument("training_data", help="Path to training data JSON file")
    parser.add_argument("--bio_data", help="Path to BIO format JSON file (optional)")
    parser.add_argument("--output_dir", default="ner_validation_results", help="Output directory")
    
    args = parser.parse_args()
    
    # Check if files exist
    if not Path(args.training_data).exists():
        print(f"❌ Training data file not found: {args.training_data}")
        return
    
    if args.bio_data and not Path(args.bio_data).exists():
        print(f"❌ BIO data file not found: {args.bio_data}")
        return
    
    # Run validation
    validator = NERTrainingValidator(args.training_data, args.bio_data)
    results = validator.run_full_validation()
    
    # Print key recommendations
    print("\n🎯 Key Recommendations:")
    for rec in results['recommendations'][:5]:
        print(f"  {rec}")

if __name__ == "__main__":
    main()