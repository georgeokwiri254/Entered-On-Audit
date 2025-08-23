"""
Entered On Report Converter
Processes the ENTERED ON sheet from Excel files and splits stay periods across months.

Key Logic:
- If a guest checks in on 30 August and checks out on 2 September:
  * 2 days in August, 1 day in September (checkout day not counted)
  * Total stays = 3 nights
- Split AMOUNT and room nights proportionally across months
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def split_stay_across_months(arrival_date, departure_date, total_amount, total_nights):
    """
    Split a stay period across months, calculating nights and amount per month.
    
    Args:
        arrival_date: Check-in date
        departure_date: Check-out date  
        total_amount: Total booking amount
        total_nights: Total nights stayed
        
    Returns:
        List of dictionaries with month splits
    """
    splits = []
    current_date = arrival_date
    
    while current_date < departure_date:
        # Calculate end of current month or departure date, whichever is earlier
        end_of_month = (current_date + relativedelta(months=1)).replace(day=1) - timedelta(days=1)
        period_end = min(end_of_month, departure_date - timedelta(days=1))  # Don't count checkout day
        
        # Calculate nights in this month
        if period_end >= current_date:
            nights_in_month = (period_end - current_date).days + 1
            
            # Calculate proportional amount for this month
            amount_for_month = (nights_in_month / total_nights) * total_amount if total_nights > 0 else 0
            
            splits.append({
                'month': current_date.strftime('%Y-%m'),
                'year': current_date.year,
                'month_num': current_date.month,
                'nights_in_month': nights_in_month,
                'amount_in_month': amount_for_month,
                'period_start': current_date,
                'period_end': period_end
            })
        
        # Move to first day of next month
        current_date = (current_date + relativedelta(months=1)).replace(day=1)
    
    return splits

def flag_long_bookings(nights):
    """Flag bookings longer than 10 days"""
    return 1 if nights > 10 else 0

def determine_season(arrival_date):
    """Determine if booking is in summer or winter season"""
    month = arrival_date.month
    # Summer: May to September (5-9), Winter: October to April (10-12, 1-4)
    if month in [5, 6, 7, 8, 9]:
        return 'Summer'
    else:
        return 'Winter'

def create_monthly_matrix(df):
    """
    Create monthly matrix format as shown in the screenshot.
    Converts split data into columns from AUG 2025 till DEC 2026.
    Creates both NIGHTS and AMOUNT columns for each month.
    """
    # Generate monthly columns from AUG 2025 till DEC 2026 (16 months)
    start_date = datetime(2025, 8, 1)  # August 2025
    end_date = datetime(2026, 12, 31)   # December 2026
    
    months_list = []
    temp_date = start_date
    
    while temp_date <= end_date:
        month_abbr = temp_date.strftime('%b').upper()
        year = temp_date.year
        # Always include year for consistency
        column_name = f"{month_abbr}{year}"
        months_list.append((column_name, temp_date))
        temp_date = (temp_date + relativedelta(months=1))
    
    # Group by booking ID (RESV ID or FULL_NAME) to aggregate monthly data
    if 'RESV ID' in df.columns:
        booking_groups = df.groupby('RESV ID')
    else:
        booking_groups = df.groupby('FULL_NAME')
    
    matrix_rows = []
    
    for booking_id, group in booking_groups:
        # Get the first row for base information
        base_row = group.iloc[0].copy()
        
        # Initialize all monthly columns to 0 (both nights and amounts)
        for month_col, _ in months_list:
            base_row[month_col] = 0  # Nights column
            base_row[f"{month_col}_AMT"] = 0  # Amount column
        
        # Fill in the actual monthly values
        for _, split_row in group.iterrows():
            split_date = split_row['PERIOD_START']
            
            # Find the matching column for this date
            for month_col, col_date in months_list:
                if (split_date.year == col_date.year and 
                    split_date.month == col_date.month):
                    base_row[month_col] = split_row['NIGHTS_IN_MONTH']
                    base_row[f"{month_col}_AMT"] = split_row['AMOUNT_IN_MONTH']
                    break
        
        matrix_rows.append(base_row)
    
    nights_columns = [col for col, _ in months_list]
    amount_columns = [f"{col}_AMT" for col, _ in months_list]
    logger.info(f"Created monthly matrix with nights columns: {nights_columns}")
    logger.info(f"Created monthly matrix with amount columns: {amount_columns}")
    return pd.DataFrame(matrix_rows)

def process_entered_on_report(file_path, output_csv_path=None):
    """
    Process the Entered On Excel file and create expanded CSV with month splits and monthly matrix.
    
    Args:
        file_path: Path to the Excel file
        output_csv_path: Path for output CSV (optional)
        
    Returns:
        Tuple of (DataFrame, csv_path)
    """
    try:
        logger.info(f"Processing Entered On report from: {file_path}")
        
        # Read the ENTERED ON sheet
        df = pd.read_excel(file_path, sheet_name='ENTERED ON')
        logger.info(f"Loaded {len(df)} records from ENTERED ON sheet")
        
        # Fill empty C_T_S_NAME values with "Brand.com" (don't rename column)
        if 'C_T_S_NAME' in df.columns:
            df['C_T_S_NAME'] = df['C_T_S_NAME'].fillna('Brand.com')
        
        # Filter out rows where Room type is 'PM'
        initial_count = len(df)
        if 'ROOM' in df.columns:
            df = df[df['ROOM'] != 'PM']
            filtered_count = len(df)
            logger.info(f"Filtered out {initial_count - filtered_count} records with Room type 'PM'")
        
        # Filter out rows where guest name contains 'room move' (case insensitive)
        room_move_count = len(df)
        if 'FULL_NAME' in df.columns:
            df = df[~df['FULL_NAME'].str.contains('room move', case=False, na=False)]
            filtered_room_move_count = len(df)
            logger.info(f"Filtered out {room_move_count - filtered_room_move_count} records with guest name containing 'room move'")
        
        # Also filter by FIRST NAME column if it exists
        if 'FIRST NAME' in df.columns:
            first_name_count = len(df)
            df = df[~df['FIRST NAME'].str.contains('room move', case=False, na=False)]
            filtered_first_name_count = len(df)
            logger.info(f"Filtered out {first_name_count - filtered_first_name_count} additional records with first name containing 'room move'")
        
        # Convert date columns to datetime then format as dd/mm/yyyy for display
        df['ARRIVAL'] = pd.to_datetime(df['ARRIVAL'])
        df['DEPARTURE'] = pd.to_datetime(df['DEPARTURE'])
        
        # Create formatted date columns for display (dd/mm/yyyy format)
        df['ARRIVAL_FORMATTED'] = df['ARRIVAL'].dt.strftime('%d/%m/%Y')
        df['DEPARTURE_FORMATTED'] = df['DEPARTURE'].dt.strftime('%d/%m/%Y')
        
        # Keep NET column as original - don't modify it
        
        # Apply 1.1x multiplier to original AMOUNT column if it exists
        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = df['AMOUNT'] * 1.1
        
        # Apply 1.1x multiplier to original ADR column (Column O) if it exists  
        if 'ADR' in df.columns:
            df['ADR'] = df['ADR'] * 1.1
        
        # Calculate additional fields (check if Season column already exists)
        if 'Season' not in df.columns:
            df['SEASON'] = df['ARRIVAL'].apply(determine_season)
        else:
            # Use existing Season column and clean it
            df['SEASON'] = df['Season'].fillna('Unknown')
        df['LONG_BOOKING_FLAG'] = df['NIGHTS'].apply(flag_long_bookings)
        
        # Create expanded dataframe with month splits
        expanded_rows = []
        
        for _, row in df.iterrows():
            arrival = row['ARRIVAL']
            departure = row['DEPARTURE']
            total_amount = row['AMOUNT'] if pd.notna(row['AMOUNT']) else 0
            total_nights = row['NIGHTS'] if pd.notna(row['NIGHTS']) else 0
            adr_multiplied = row['ADR'] if pd.notna(row['ADR']) else 0  # ADR is already multiplied by 1.1
            
            # Get month splits for this booking
            month_splits = split_stay_across_months(arrival, departure, total_amount, total_nights)
            
            # Create a row for each month split
            for split in month_splits:
                new_row = row.copy()
                
                # Add split-specific columns
                new_row['SPLIT_MONTH'] = split['month']
                new_row['SPLIT_YEAR'] = split['year'] 
                new_row['SPLIT_MONTH_NUM'] = split['month_num']
                new_row['NIGHTS_IN_MONTH'] = split['nights_in_month']
                
                # Calculate amount for this month: ADR × nights_in_month (ADR already multiplied by 1.1)
                new_row['AMOUNT_IN_MONTH'] = adr_multiplied * split['nights_in_month']
                
                new_row['PERIOD_START'] = split['period_start']
                new_row['PERIOD_END'] = split['period_end']
                
                # ADR for this month is the same as the booking ADR (already multiplied by 1.1)
                new_row['ADR_IN_MONTH'] = adr_multiplied
                
                expanded_rows.append(new_row)
        
        # Create expanded DataFrame
        expanded_df = pd.DataFrame(expanded_rows)
        
        # Add derived columns for analysis (avoid duplicates)
        if 'BOOKING_LEAD_TIME' not in expanded_df.columns:
            expanded_df['BOOKING_LEAD_TIME'] = expanded_df.get('Booking Lead Time ', 0)
        if 'EVENTS_DATES' not in expanded_df.columns:
            expanded_df['EVENTS_DATES'] = expanded_df.get('Events Dates ', '')
        
        # Clean company name for analysis
        if 'C_T_S_NAME' in expanded_df.columns:
            expanded_df['COMPANY_CLEAN'] = expanded_df['C_T_S_NAME'].fillna('Brand.com')
        else:
            expanded_df['COMPANY_CLEAN'] = 'Brand.com'
        
        # Drop duplicate columns if they exist
        columns_to_drop = ['Season', 'Booking Lead Time ', 'Events Dates ']
        for col in columns_to_drop:
            if col in expanded_df.columns:
                expanded_df = expanded_df.drop(columns=[col])
        
        logger.info(f"Expanded to {len(expanded_df)} records after month splitting")
        
        # Create monthly matrix format
        matrix_df = create_monthly_matrix(expanded_df)
        logger.info(f"Created monthly matrix with {len(matrix_df)} bookings")
        
        # Remove split-related columns since we now have monthly matrix columns
        split_columns_to_remove = ['SPLIT_MONTH', 'SPLIT_YEAR', 'SPLIT_MONTH_NUM', 'NIGHTS_IN_MONTH', 'AMOUNT_IN_MONTH', 'PERIOD_START', 'PERIOD_END', 'ADR_IN_MONTH']
        for col in split_columns_to_remove:
            if col in matrix_df.columns:
                matrix_df = matrix_df.drop(columns=[col])
        logger.info(f"Removed split columns, now have {len(matrix_df.columns)} columns")
        
        # Set output path if not provided
        if output_csv_path is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_csv_path = f"data/processed/entered_on_{timestamp}.csv"
        
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)
        
        # Save the monthly matrix format (this is what user wants to see)
        matrix_df.to_csv(output_csv_path, index=False)
        logger.info(f"Saved monthly matrix data to: {output_csv_path}")
        
        # Always try to save canonical version - use multiple strategies
        canonical_path = "data/processed/entered_on.csv"
        temp_canonical_path = "data/processed/entered_on_temp.csv"
        
        success = False
        
        # Strategy 1: Direct overwrite (works if file is not locked)
        try:
            matrix_df.to_csv(canonical_path, index=False)
            logger.info(f"Saved canonical version to: {canonical_path}")
            success = True
        except PermissionError:
            logger.warning(f"Direct overwrite failed - file may be in use")
        
        # Strategy 2: Atomic write with temp file + rename (if direct failed)
        if not success:
            try:
                # Write to temporary file first
                matrix_df.to_csv(temp_canonical_path, index=False)
                
                # Try to remove old file
                if os.path.exists(canonical_path):
                    try:
                        os.remove(canonical_path)
                    except PermissionError:
                        # Try a few more times with small delays
                        import time
                        for attempt in range(3):
                            time.sleep(0.1)  # Wait 100ms
                            try:
                                os.remove(canonical_path)
                                break
                            except PermissionError:
                                continue
                        else:
                            # Still can't remove - clean up and use timestamped version
                            logger.warning(f"Could not replace canonical file after multiple attempts")
                            os.remove(temp_canonical_path)
                            return matrix_df, output_csv_path
                
                # Rename temp file to canonical name
                os.rename(temp_canonical_path, canonical_path)
                logger.info(f"Saved canonical version to: {canonical_path} (using atomic write)")
                success = True
                
            except Exception as e:
                # Clean up temp file if it exists
                if os.path.exists(temp_canonical_path):
                    try:
                        os.remove(temp_canonical_path)
                    except:
                        pass
                logger.warning(f"Atomic write failed: {e}")
        
        if success:
            return matrix_df, canonical_path
        else:
            logger.warning(f"Could not save canonical version, using timestamped file: {output_csv_path}")
            return matrix_df, output_csv_path
        
    except Exception as e:
        logger.error(f"Error processing Entered On report: {e}")
        raise

def get_summary_stats(df):
    """Get summary statistics for the processed data"""
    stats = {
        'total_bookings': len(df['RESV ID'].unique()) if 'RESV ID' in df.columns else len(df),
        'total_amount': df['AMOUNT_IN_MONTH'].sum(),
        'total_room_nights': df['NIGHTS_IN_MONTH'].sum(),
        'average_adr': df['ADR_IN_MONTH'].mean(),
        'long_bookings': df['LONG_BOOKING_FLAG'].sum(),
        'months_covered': df['SPLIT_MONTH'].nunique(),
        'companies_count': df['COMPANY_CLEAN'].nunique(),
        'seasons_breakdown': df.groupby('SEASON')['NIGHTS_IN_MONTH'].sum().to_dict()
    }
    return stats

if __name__ == "__main__":
    # Test the converter
    file_path = "16-08-2025 Entered On.xlsm"
    if os.path.exists(file_path):
        df, csv_path = process_entered_on_report(file_path)
        stats = get_summary_stats(df)
        print("Summary Statistics:")
        for key, value in stats.items():
            print(f"  {key}: {value}")
    else:
        print(f"File not found: {file_path}")