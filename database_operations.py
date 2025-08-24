"""
Database Operations Module for Entered On Audit System
Handles SQLite database operations for reservations, emails, audit results, and logging.
"""

import sqlite3
import pandas as pd
import json
import uuid
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any
import logging
from contextlib import contextmanager

logger = logging.getLogger(__name__)

class AuditDatabase:
    """Main database class for the Entered On Audit System"""
    
    def __init__(self, db_path: str = "data/audit_database.db"):
        self.db_path = db_path
        self.init_database()
    
    @contextmanager
    def get_connection(self):
        """Context manager for database connections"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # Enable column access by name
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def _migrate_schema(self, conn):
        """Handle schema migrations for existing databases"""
        try:
            # Add NET column to reservations_raw table if it doesn't exist
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(reservations_raw)")
            columns = [col[1] for col in cursor.fetchall()]
            
            if 'net' not in columns:
                logger.info("Adding 'net' column to reservations_raw table")
                conn.execute("ALTER TABLE reservations_raw ADD COLUMN net REAL DEFAULT 0.0")
            
            # Add mail_c_t_s_name and mail_net columns to reservations_email table
            cursor.execute("PRAGMA table_info(reservations_email)")
            columns = [col[1] for col in cursor.fetchall()]
            
            if 'mail_c_t_s_name' not in columns:
                logger.info("Adding 'mail_c_t_s_name' column to reservations_email table")
                conn.execute("ALTER TABLE reservations_email ADD COLUMN mail_c_t_s_name TEXT DEFAULT ''")
            
            if 'mail_net' not in columns:
                logger.info("Adding 'mail_net' column to reservations_email table")
                conn.execute("ALTER TABLE reservations_email ADD COLUMN mail_net REAL DEFAULT 0.0")
            
            # Add mail_c_t_s_name and mail_net columns to reservations_audit table
            cursor.execute("PRAGMA table_info(reservations_audit)")
            columns = [col[1] for col in cursor.fetchall()]
            
            if 'mail_c_t_s_name' not in columns:
                logger.info("Adding 'mail_c_t_s_name' column to reservations_audit table")
                conn.execute("ALTER TABLE reservations_audit ADD COLUMN mail_c_t_s_name TEXT DEFAULT ''")
            
            if 'mail_net' not in columns:
                logger.info("Adding 'mail_net' column to reservations_audit table")
                conn.execute("ALTER TABLE reservations_audit ADD COLUMN mail_net REAL DEFAULT 0.0")
            
        except Exception as e:
            logger.warning(f"Schema migration failed: {e}")
    
    def init_database(self):
        """Initialize database with all required tables"""
        with self.get_connection() as conn:
            # Run schema migrations first
            self._migrate_schema(conn)
            # Create audit_log table (run tracking)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS audit_log (
                    run_id TEXT PRIMARY KEY,
                    run_timestamp TEXT NOT NULL,
                    excel_file_processed TEXT,
                    reservations_loaded_count INTEGER DEFAULT 0,
                    emails_found_count INTEGER DEFAULT 0,
                    pdf_extractions_count INTEGER DEFAULT 0,
                    audit_pass_count INTEGER DEFAULT 0,
                    audit_fail_count INTEGER DEFAULT 0,
                    errors_encountered TEXT DEFAULT '[]',
                    execution_time_seconds REAL DEFAULT 0.0,
                    status TEXT DEFAULT 'RUNNING',
                    notes TEXT DEFAULT ''
                )
            """)
            
            # Create reservations_raw table (Excel data)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS reservations_raw (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT NOT NULL,
                    full_name TEXT,
                    first_name TEXT,
                    arrival TEXT,
                    departure TEXT,
                    nights INTEGER,
                    persons INTEGER,
                    room TEXT,
                    rate_code TEXT,
                    c_t_s_name TEXT,
                    net REAL,
                    net_total REAL,
                    amount REAL,
                    adr REAL,
                    season TEXT,
                    long_booking_flag INTEGER DEFAULT 0,
                    company_clean TEXT,
                    booking_lead_time INTEGER,
                    events_dates TEXT,
                    resv_id TEXT,
                    created_timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
                    raw_data TEXT,  -- JSON storage for additional columns
                    FOREIGN KEY (run_id) REFERENCES audit_log(run_id)
                )
            """)
            
            # Create reservations_email table (Email extraction)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS reservations_email (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT NOT NULL,
                    guest_name TEXT,
                    email_subject TEXT,
                    email_sender TEXT,
                    email_received_time TEXT,
                    folder_name TEXT,
                    mail_first_name TEXT,
                    mail_arrival TEXT,
                    mail_departure TEXT,
                    mail_nights INTEGER,
                    mail_persons INTEGER,
                    mail_room TEXT,
                    mail_rate_code TEXT,
                    mail_c_t_s TEXT,
                    mail_c_t_s_name TEXT,
                    mail_net REAL,
                    mail_net_total REAL,
                    mail_total REAL,
                    mail_tdf REAL,
                    mail_adr REAL,
                    mail_amount REAL,
                    pdf_attachment_count INTEGER DEFAULT 0,
                    extraction_timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
                    raw_email_data TEXT,  -- JSON storage for full email data
                    FOREIGN KEY (run_id) REFERENCES audit_log(run_id)
                )
            """)
            
            # Create reservations_audit table (Final results)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS reservations_audit (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT NOT NULL,
                    -- Original reservation data
                    full_name TEXT,
                    first_name TEXT,
                    arrival TEXT,
                    departure TEXT,
                    nights INTEGER,
                    persons INTEGER,
                    room TEXT,
                    rate_code TEXT,
                    c_t_s_name TEXT,
                    net REAL,
                    net_total REAL,
                    amount REAL,
                    adr REAL,
                    season TEXT,
                    long_booking_flag INTEGER DEFAULT 0,
                    company_clean TEXT,
                    -- Email extracted data (Mail_ prefix)
                    mail_first_name TEXT,
                    mail_arrival TEXT,
                    mail_departure TEXT,
                    mail_nights INTEGER,
                    mail_persons INTEGER,
                    mail_room TEXT,
                    mail_rate_code TEXT,
                    mail_c_t_s TEXT,
                    mail_c_t_s_name TEXT,
                    mail_net REAL,
                    mail_net_total REAL,
                    mail_total REAL,
                    mail_tdf REAL,
                    mail_adr REAL,
                    mail_amount REAL,
                    -- Audit results
                    audit_status TEXT DEFAULT 'PENDING',
                    audit_issues TEXT DEFAULT '',
                    fields_matching INTEGER DEFAULT 0,
                    total_email_fields INTEGER DEFAULT 0,
                    match_percentage REAL DEFAULT 0.0,
                    email_vs_data_status TEXT DEFAULT 'N/A',
                    audit_timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
                    raw_audit_data TEXT,  -- JSON storage for additional audit info
                    FOREIGN KEY (run_id) REFERENCES audit_log(run_id)
                )
            """)
            
            # Create indices for better query performance
            conn.execute("CREATE INDEX IF NOT EXISTS idx_raw_run_id ON reservations_raw(run_id)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_raw_guest ON reservations_raw(full_name)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_email_run_id ON reservations_email(run_id)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_email_guest ON reservations_email(guest_name)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_run_id ON reservations_audit(run_id)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_status ON reservations_audit(audit_status)")
            
        logger.info(f"Database initialized at {self.db_path}")
    
    def start_run(self, excel_file: str = None) -> str:
        """Start a new audit run and return run_id"""
        run_id = f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{str(uuid.uuid4())[:8]}"
        
        with self.get_connection() as conn:
            conn.execute("""
                INSERT INTO audit_log (
                    run_id, run_timestamp, excel_file_processed, status
                ) VALUES (?, ?, ?, ?)
            """, (run_id, datetime.now().isoformat(), excel_file, 'RUNNING'))
        
        logger.info(f"Started audit run: {run_id}")
        return run_id
    
    def update_run_stats(self, run_id: str, stats: Dict[str, Any]):
        """Update statistics for a run"""
        with self.get_connection() as conn:
            update_fields = []
            values = []
            
            for key, value in stats.items():
                if key in ['reservations_loaded_count', 'emails_found_count', 'pdf_extractions_count',
                          'audit_pass_count', 'audit_fail_count', 'execution_time_seconds', 'status', 'notes']:
                    update_fields.append(f"{key} = ?")
                    values.append(value)
            
            if update_fields:
                values.append(run_id)
                query = f"UPDATE audit_log SET {', '.join(update_fields)} WHERE run_id = ?"
                conn.execute(query, values)
                logger.debug(f"Updated run stats for {run_id}: {stats}")
    
    def log_error(self, run_id: str, error: str, context: str = ""):
        """Add an error to the run log"""
        with self.get_connection() as conn:
            # Get current errors
            result = conn.execute("SELECT errors_encountered FROM audit_log WHERE run_id = ?", (run_id,))
            row = result.fetchone()
            
            if row:
                current_errors = json.loads(row['errors_encountered'] or '[]')
                current_errors.append({
                    'timestamp': datetime.now().isoformat(),
                    'error': error,
                    'context': context
                })
                
                conn.execute("""
                    UPDATE audit_log SET errors_encountered = ? WHERE run_id = ?
                """, (json.dumps(current_errors), run_id))
                
                logger.error(f"Run {run_id} - {context}: {error}")
    
    def save_raw_reservations(self, df: pd.DataFrame, run_id: str) -> int:
        """Save raw reservation data from Excel processing"""
        try:
            with self.get_connection() as conn:
                # Prepare data for insertion
                records = []
                for _, row in df.iterrows():
                    # Extract standard fields
                    record = {
                        'run_id': run_id,
                        'full_name': row.get('FULL_NAME', ''),
                        'first_name': row.get('FIRST_NAME', ''),
                        'arrival': row.get('ARRIVAL', ''),
                        'departure': row.get('DEPARTURE', ''),
                        'nights': row.get('NIGHTS', 0),
                        'persons': row.get('PERSONS', 0),
                        'room': row.get('ROOM', ''),
                        'rate_code': row.get('RATE_CODE', ''),
                        'c_t_s_name': row.get('C_T_S_NAME', ''),
                        'net': row.get('NET', 0.0),
                        'net_total': row.get('NET_TOTAL', 0.0),
                        'amount': row.get('AMOUNT', 0.0),
                        'adr': row.get('ADR', 0.0),
                        'season': row.get('SEASON', ''),
                        'long_booking_flag': row.get('LONG_BOOKING_FLAG', 0),
                        'company_clean': row.get('COMPANY_CLEAN', ''),
                        'booking_lead_time': row.get('BOOKING_LEAD_TIME', 0),
                        'events_dates': row.get('EVENTS_DATES', ''),
                        'resv_id': row.get('RESV ID', ''),
                        'raw_data': json.dumps(self._serialize_pandas_row(row))  # Store full row as JSON
                    }
                    records.append(record)
                
                # Insert records
                placeholders = ', '.join(['?' for _ in record.keys()])
                columns = ', '.join(record.keys())
                
                for record in records:
                    conn.execute(f"""
                        INSERT INTO reservations_raw ({columns})
                        VALUES ({placeholders})
                    """, list(record.values()))
                
            count = len(records)
            self.update_run_stats(run_id, {'reservations_loaded_count': count})
            logger.info(f"Saved {count} raw reservations for run {run_id}")
            return count
            
        except Exception as e:
            self.log_error(run_id, str(e), "save_raw_reservations")
            raise
    
    def save_email_extraction(self, email_results: List[Dict], run_id: str) -> int:
        """Save email extraction results"""
        try:
            with self.get_connection() as conn:
                count = 0
                
                for result in email_results:
                    guest_name = result['reservation_data'].get('FULL_NAME', '')
                    
                    for email in result.get('matching_emails', []):
                        extracted_data = email.get('extracted_data', {})
                        
                        record = {
                            'run_id': run_id,
                            'guest_name': guest_name,
                            'email_subject': email.get('subject', ''),
                            'email_sender': email.get('sender', ''),
                            'email_received_time': str(email.get('received_time', '')),
                            'folder_name': email.get('folder', ''),
                            'mail_first_name': extracted_data.get('FIRST_NAME', ''),
                            'mail_arrival': extracted_data.get('ARRIVAL', ''),
                            'mail_departure': extracted_data.get('DEPARTURE', ''),
                            'mail_nights': extracted_data.get('NIGHTS', 0) if extracted_data.get('NIGHTS') != 'N/A' else 0,
                            'mail_persons': extracted_data.get('PERSONS', 0) if extracted_data.get('PERSONS') != 'N/A' else 0,
                            'mail_room': extracted_data.get('ROOM', ''),
                            'mail_rate_code': extracted_data.get('RATE_CODE', ''),
                            'mail_c_t_s': extracted_data.get('C_T_S', ''),
                            'mail_c_t_s_name': extracted_data.get('C_T_S_NAME', ''),
                            'mail_net': self._parse_float(extracted_data.get('NET', 0)),
                            'mail_net_total': self._parse_float(extracted_data.get('NET_TOTAL', 0)),
                            'mail_total': self._parse_float(extracted_data.get('TOTAL', 0)),
                            'mail_tdf': self._parse_float(extracted_data.get('TDF', 0)),
                            'mail_adr': self._parse_float(extracted_data.get('ADR', 0)),
                            'mail_amount': self._parse_float(extracted_data.get('AMOUNT', 0)),
                            'pdf_attachment_count': len([att for att in email.get('attachments', []) if att.get('filename', '').lower().endswith('.pdf')]),
                            'raw_email_data': json.dumps(self._serialize_email_data(email))
                        }
                        
                        placeholders = ', '.join(['?' for _ in record.keys()])
                        columns = ', '.join(record.keys())
                        
                        conn.execute(f"""
                            INSERT INTO reservations_email ({columns})
                            VALUES ({placeholders})
                        """, list(record.values()))
                        
                        count += 1
                
            self.update_run_stats(run_id, {
                'emails_found_count': count,
                'pdf_extractions_count': sum(1 for result in email_results if result['has_pdf_data'])
            })
            logger.info(f"Saved {count} email extractions for run {run_id}")
            return count
            
        except Exception as e:
            self.log_error(run_id, str(e), "save_email_extraction")
            raise
    
    def save_audit_results(self, audit_df: pd.DataFrame, run_id: str) -> int:
        """Save final audit results"""
        try:
            with self.get_connection() as conn:
                records = []
                pass_count = 0
                fail_count = 0
                
                for _, row in audit_df.iterrows():
                    # Track pass/fail counts
                    if row.get('audit_status') == 'PASS':
                        pass_count += 1
                    else:
                        fail_count += 1
                    
                    record = {
                        'run_id': run_id,
                        # Original data
                        'full_name': row.get('FULL_NAME', ''),
                        'first_name': row.get('FIRST_NAME', ''),
                        'arrival': row.get('ARRIVAL', ''),
                        'departure': row.get('DEPARTURE', ''),
                        'nights': row.get('NIGHTS', 0),
                        'persons': row.get('PERSONS', 0),
                        'room': row.get('ROOM', ''),
                        'rate_code': row.get('RATE_CODE', ''),
                        'c_t_s_name': row.get('C_T_S_NAME', ''),
                        'net': row.get('NET', 0.0),
                        'net_total': row.get('NET_TOTAL', 0.0),
                        'amount': row.get('AMOUNT', 0.0),
                        'adr': row.get('ADR', 0.0),
                        'season': row.get('SEASON', ''),
                        'long_booking_flag': row.get('LONG_BOOKING_FLAG', 0),
                        'company_clean': row.get('COMPANY_CLEAN', ''),
                        # Mail data
                        'mail_first_name': row.get('Mail_FIRST_NAME', ''),
                        'mail_arrival': row.get('Mail_ARRIVAL', ''),
                        'mail_departure': row.get('Mail_DEPARTURE', ''),
                        'mail_nights': row.get('Mail_NIGHTS', 0),
                        'mail_persons': row.get('Mail_PERSONS', 0),
                        'mail_room': row.get('Mail_ROOM', ''),
                        'mail_rate_code': row.get('Mail_RATE_CODE', ''),
                        'mail_c_t_s': row.get('Mail_C_T_S', ''),
                        'mail_c_t_s_name': row.get('Mail_C_T_S_NAME', ''),
                        'mail_net': row.get('Mail_NET', 0.0),
                        'mail_net_total': row.get('Mail_NET_TOTAL', 0.0),
                        'mail_total': row.get('Mail_TOTAL', 0.0),
                        'mail_tdf': row.get('Mail_TDF', 0.0),
                        'mail_adr': row.get('Mail_ADR', 0.0),
                        'mail_amount': row.get('Mail_AMOUNT', 0.0),
                        # Audit results
                        'audit_status': row.get('audit_status', 'PENDING'),
                        'audit_issues': row.get('audit_issues', ''),
                        'fields_matching': row.get('fields_matching', 0),
                        'total_email_fields': row.get('total_email_fields', 0),
                        'match_percentage': row.get('match_percentage', 0.0),
                        'email_vs_data_status': row.get('email_vs_data_status', 'N/A'),
                        'raw_audit_data': json.dumps(self._serialize_pandas_row(row))
                    }
                    records.append(record)
                
                # Insert records
                for record in records:
                    placeholders = ', '.join(['?' for _ in record.keys()])
                    columns = ', '.join(record.keys())
                    
                    conn.execute(f"""
                        INSERT INTO reservations_audit ({columns})
                        VALUES ({placeholders})
                    """, list(record.values()))
                
            # Update run statistics
            self.update_run_stats(run_id, {
                'audit_pass_count': pass_count,
                'audit_fail_count': fail_count,
                'status': 'COMPLETED'
            })
            
            count = len(records)
            logger.info(f"Saved {count} audit results for run {run_id}")
            return count
            
        except Exception as e:
            self.log_error(run_id, str(e), "save_audit_results")
            self.update_run_stats(run_id, {'status': 'FAILED'})
            raise
    
    def _parse_float(self, value) -> float:
        """Helper to parse float values, handling strings and N/A"""
        if value == 'N/A' or value is None:
            return 0.0
        try:
            if isinstance(value, str):
                # Remove currency symbols and commas
                cleaned = value.replace('AED', '').replace(',', '').strip()
                return float(cleaned)
            return float(value)
        except (ValueError, TypeError):
            return 0.0
    
    def _serialize_pandas_row(self, row) -> dict:
        """Helper to serialize pandas row to JSON-compatible dict, handling Timestamps"""
        result = {}
        for key, value in row.to_dict().items():
            if pd.isna(value):
                result[key] = None
            elif hasattr(value, 'isoformat'):  # datetime/Timestamp objects
                result[key] = value.isoformat()
            elif isinstance(value, (pd.Timestamp, datetime)):
                result[key] = str(value)
            else:
                result[key] = value
        return result
    
    def _serialize_email_data(self, email_data) -> dict:
        """Helper to serialize email data to JSON-compatible dict, handling datetime objects"""
        result = {}
        for key, value in email_data.items():
            if hasattr(value, 'isoformat'):  # datetime/Timestamp objects
                result[key] = value.isoformat()
            elif isinstance(value, (pd.Timestamp, datetime)):
                result[key] = str(value)
            elif isinstance(value, dict):
                result[key] = self._serialize_email_data(value)  # Recursive for nested dicts
            elif isinstance(value, list):
                result[key] = [self._serialize_email_data(item) if isinstance(item, dict) else item for item in value]
            else:
                result[key] = value
        return result
    
    def get_recent_runs(self, limit: int = 10) -> pd.DataFrame:
        """Get recent audit runs"""
        with self.get_connection() as conn:
            query = """
                SELECT run_id, run_timestamp, excel_file_processed,
                       reservations_loaded_count, emails_found_count, pdf_extractions_count,
                       audit_pass_count, audit_fail_count, status, execution_time_seconds
                FROM audit_log 
                ORDER BY run_timestamp DESC 
                LIMIT ?
            """
            return pd.read_sql_query(query, conn, params=(limit,))
    
    def get_run_errors(self, run_id: str) -> List[Dict]:
        """Get errors for a specific run"""
        with self.get_connection() as conn:
            result = conn.execute("SELECT errors_encountered FROM audit_log WHERE run_id = ?", (run_id,))
            row = result.fetchone()
            if row and row['errors_encountered']:
                return json.loads(row['errors_encountered'])
            return []
    
    def get_audit_results(self, run_id: str = None, filters: Dict = None) -> pd.DataFrame:
        """Get audit results with optional filtering"""
        with self.get_connection() as conn:
            query = "SELECT * FROM reservations_audit"
            params = []
            conditions = []
            
            if run_id:
                conditions.append("run_id = ?")
                params.append(run_id)
            
            if filters:
                if filters.get('audit_status'):
                    conditions.append("audit_status = ?")
                    params.append(filters['audit_status'])
                
                if filters.get('email_vs_data_status'):
                    conditions.append("email_vs_data_status = ?")
                    params.append(filters['email_vs_data_status'])
                
                if filters.get('season'):
                    conditions.append("season = ?")
                    params.append(filters['season'])
                
                if filters.get('company_clean'):
                    conditions.append("company_clean = ?")
                    params.append(filters['company_clean'])
            
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            query += " ORDER BY audit_timestamp DESC"
            
            return pd.read_sql_query(query, conn, params=params)
    
    def get_summary_stats(self, run_id: str = None) -> Dict:
        """Get summary statistics"""
        with self.get_connection() as conn:
            if run_id:
                # Stats for specific run
                audit_query = "SELECT * FROM reservations_audit WHERE run_id = ?"
                audit_df = pd.read_sql_query(audit_query, conn, params=(run_id,))
                
                if audit_df.empty:
                    return {}
                
                return {
                    'total_reservations': len(audit_df),
                    'audit_pass': len(audit_df[audit_df['audit_status'] == 'PASS']),
                    'audit_fail': len(audit_df[audit_df['audit_status'] == 'FAIL']),
                    'email_match_pass': len(audit_df[audit_df['email_vs_data_status'] == 'PASS']),
                    'with_email_data': len(audit_df[audit_df['total_email_fields'] > 0]),
                    'avg_match_percentage': audit_df['match_percentage'].mean(),
                    'total_amount': audit_df['amount'].sum(),
                    'total_nights': audit_df['nights'].sum()
                }
            else:
                # Overall stats
                run_query = "SELECT COUNT(*) as total_runs FROM audit_log WHERE status = 'COMPLETED'"
                audit_query = "SELECT COUNT(*) as total_audits FROM reservations_audit"
                
                run_count = conn.execute(run_query).fetchone()['total_runs']
                audit_count = conn.execute(audit_query).fetchone()['total_audits']
                
                return {
                    'total_runs': run_count,
                    'total_audits': audit_count
                }
    
    def export_data(self, table: str, run_id: str = None, filters: Dict = None) -> pd.DataFrame:
        """Export data from any table with optional filtering"""
        if table not in ['reservations_raw', 'reservations_email', 'reservations_audit', 'audit_log']:
            raise ValueError(f"Invalid table name: {table}")
        
        with self.get_connection() as conn:
            query = f"SELECT * FROM {table}"
            params = []
            conditions = []
            
            if run_id and table != 'audit_log':
                conditions.append("run_id = ?")
                params.append(run_id)
            
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            return pd.read_sql_query(query, conn, params=params)
    
    def cleanup_old_runs(self, days_to_keep: int = 30):
        """Clean up old runs (optional maintenance function)"""
        cutoff_date = datetime.now() - pd.Timedelta(days=days_to_keep)
        
        with self.get_connection() as conn:
            # Get old run IDs
            old_runs = conn.execute("""
                SELECT run_id FROM audit_log 
                WHERE run_timestamp < ? AND status != 'RUNNING'
            """, (cutoff_date.isoformat(),)).fetchall()
            
            if old_runs:
                old_run_ids = [row['run_id'] for row in old_runs]
                placeholders = ','.join(['?' for _ in old_run_ids])
                
                # Delete from all tables
                for table in ['reservations_audit', 'reservations_email', 'reservations_raw', 'audit_log']:
                    conn.execute(f"DELETE FROM {table} WHERE run_id IN ({placeholders})", old_run_ids)
                
                logger.info(f"Cleaned up {len(old_run_ids)} old runs")
                return len(old_run_ids)
            
            return 0