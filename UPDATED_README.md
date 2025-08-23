# Entered On Audit System - Updated with Outlook Integration

## 🔄 Recent Updates

✅ **Replaced IMAP with win32com.client** for direct Outlook access  
✅ **Added AED currency handling** with USD conversion  
✅ **Reservation-based email search** - iterates through each reservation  
✅ **Local Outlook integration** - no email credentials needed  
✅ **Enhanced PDF extraction** with currency-aware regex patterns  

## 🏨 Features

### 📧 Tab 1: Email Extraction Results
- **Per-Reservation Email Search**: Searches Outlook emails for each guest reservation
- **Smart Matching**: Finds emails by guest name and date proximity
- **PDF Data Extraction**: Extracts reservation details from PDF attachments
- **AED/USD Currency Display**: Shows amounts in both currencies
- **Export Results**: Download comprehensive email extraction report

### 📊 Tab 2: Converted Data
- **Full Entered On Sheet Display**: Shows complete processed dataset
- **AED Currency Handling**: Properly displays amounts in AED
- **Monthly Splits**: Reservations split across months with proportional amounts
- **Advanced Filtering**: Filter by Season, Company, Room Type
- **Export Functionality**: Download processed data as CSV

### 🔍 Tab 3: Audit Results
- **Validation Checks**: Verify nights calculation, amount consistency, required fields
- **Pass/Fail Analysis**: Clear status indicators and detailed issue reporting
- **AED Amount Validation**: Check NET_TOTAL vs TDF in proper currency
- **Comprehensive Reporting**: Detailed failure analysis with corrective actions

## 🚀 How to Run

### Prerequisites
- **Windows OS** (required for Outlook COM access)
- **Microsoft Outlook** installed and configured
- **Python 3.8+** with required packages

### Installation
1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Launch App:**
   ```bash
   streamlit run streamlit_app.py
   ```
   or double-click `launch_app.bat`

## 📧 Outlook Integration

### How It Works
1. **Direct COM Access**: Uses win32com.client to access local Outlook
2. **No Credentials Needed**: Works with your existing Outlook installation
3. **Inbox Search**: Searches emails in configurable date range (default: 7 days)
4. **Smart Matching**: Finds emails containing guest names from reservations

### Email Search Process
For each reservation in the Excel file:
1. Extract guest name and reservation details
2. Search Outlook inbox for emails containing guest name
3. Process PDF attachments from matching emails
4. Extract reservation data using regex patterns
5. Merge email data with Excel reservation data

## 💱 Currency Handling

### AED Support
- **Source Data**: Excel amounts are in AED
- **Regex Patterns**: Updated to handle "AED 1,234.56" format
- **Conversion**: Automatic AED to USD conversion (configurable rate)
- **Display Options**: Show AED only, USD only, or both

### Exchange Rate
- Default: 1 AED = 0.27 USD
- Configurable in sidebar
- Applied to extracted PDF amounts

## 📊 Data Processing

### Excel Processing
1. **Load ENTERED ON sheet** from uploaded .xlsm file
2. **Apply filters**: Remove PM rooms, "room move" entries
3. **Calculate splits**: Divide stays across months
4. **Apply multipliers**: 1.1x to AMOUNT and ADR columns
5. **Add derived fields**: Season, booking flags, company clean names

### Email Processing
1. **Connect to Outlook** using COM interface
2. **Iterate reservations** and search for matching emails
3. **Extract PDF text** from email attachments
4. **Parse reservation fields** using enhanced regex
5. **Merge data** with Excel reservation information

## 🔍 Audit Checks

### Validation Rules
1. **Nights Calculation**: (Departure - Arrival) = Nights
2. **Amount Consistency**: NET_TOTAL ≥ TDF (in AED)
3. **Person Count**: Persons > 0
4. **Required Fields**: Name, Arrival, Departure, Nights not empty
5. **Data Format**: Valid dates and numeric amounts

### Results
- **PASS**: All validations successful
- **FAIL**: One or more validation failures
- **Detailed Issues**: Specific problems with suggested fixes

## 📁 File Structure

```
Entered On Audit/
├── streamlit_app.py          # Main Streamlit application
├── entered_on_converter.py   # Excel processing logic
├── requirements.txt          # Python dependencies
├── launch_app.bat           # Windows launcher
├── RUN_APP.md              # Detailed instructions
└── UPDATED_README.md       # This file
```

## 🛠 Troubleshooting

### Common Issues

**"Could not connect to Outlook"**
- Ensure Outlook is installed and configured
- Try opening Outlook manually first
- Check Windows permissions

**"No matching emails found"**
- Verify guest names are spelled correctly
- Increase search date range
- Check Outlook inbox has recent emails

**"PDF extraction failed"**
- Ensure PDFs contain text (not scanned images)
- Check file permissions
- Verify PDF is not password protected

**"Currency conversion error"**
- Check AED amounts are numeric
- Verify exchange rate is positive
- Look for special characters in amounts

## 📈 Performance

- **Email Search**: ~1-2 seconds per reservation
- **PDF Processing**: ~3-5 seconds per PDF attachment
- **Excel Processing**: ~10-50ms per row
- **Recommended**: Process in batches of 100-500 reservations

## 🔐 Security & Privacy

- **Local Processing**: All data stays on your machine
- **No Cloud Access**: Direct Outlook integration only
- **Temporary Files**: PDF attachments cleaned up automatically
- **No Credentials**: Uses existing Outlook authentication

## 📞 Support

For issues or questions:
1. Check troubleshooting section above
2. Verify all prerequisites are met
3. Review console output for error details
4. Ensure Outlook and Python environments are properly configured