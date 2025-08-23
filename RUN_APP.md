# How to Run the Entered On Audit Streamlit App

## Quick Start

1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the App:**
   ```bash
   streamlit run streamlit_app.py
   ```

3. **Open in Browser:**
   - The app will automatically open at `http://localhost:8501`

## App Features

### 📧 Tab 1: Email Extraction Results
- Configure IMAP email settings in the sidebar
- Fetch emails from the last 2 days (configurable)
- Extract reservation data from PDF attachments
- View extracted fields like guest name, arrival, departure, etc.

### 📊 Tab 2: Converted Data  
- Upload your `.xlsm` Entered On report file
- View the complete processed dataset
- Filter by Season, Company, Room Type
- Download processed data as CSV
- Shows all records with monthly splits and calculations

### 🔍 Tab 3: Audit Results
- Run validation checks on processed data
- Verify nights calculation (departure - arrival)
- Check NET_TOTAL vs TDF values
- Validate person counts and required fields
- View detailed failure reasons
- Download audit results

## Setup Instructions

### Email Configuration
1. Use your email provider's IMAP settings:
   - **Gmail:** `imap.gmail.com` (port 993)
   - **Outlook:** `outlook.office365.com` (port 993)
2. For Gmail, you may need to use an App Password instead of your regular password
3. Enter credentials in the sidebar when the app is running

### File Upload
1. Click "Browse files" in the sidebar
2. Select your `.xlsm` Entered On report
3. The app will automatically process and display summary metrics

## Workflow

1. **Upload Excel File** → Processes and shows converted data in Tab 2
2. **Configure Email** → Set IMAP settings in sidebar  
3. **Fetch Emails** → Extracts data from PDF attachments in Tab 1
4. **Run Audit** → Validates data and shows results in Tab 3
5. **Download Results** → Export processed data and audit reports

## Data Processing

The app integrates with `entered_on_converter.py` to:
- Split stays across multiple months
- Apply 1.1x multiplier to amounts and ADR
- Filter out PM room types and "room move" entries
- Calculate seasonal flags and long booking indicators
- Create monthly matrix format with night/amount columns

## Troubleshooting

- **Email Connection Issues:** Check IMAP settings and credentials
- **PDF Extraction Issues:** Ensure PDFs contain readable text (not scanned images)
- **File Upload Problems:** Verify file is valid Excel format (.xlsm/.xlsx)
- **Missing Data:** Check that required columns exist in Excel file

## Security

- Email credentials are not stored persistently
- PDF files are processed in memory only
- No data is transmitted outside your local environment