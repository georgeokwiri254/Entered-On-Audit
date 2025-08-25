# Entered On Audit

* **Email Access** → IMAP (direct to mailbox reservations.gmhd@millenniumhotels.com).
* **Email Attachments** → PDFs (extract reservations using regex).
* **Comparison & Audit** → merge with Excel data, fill missing fields, mark `N/A` if not found.

Here’s the **10-point workflow with Streamlit + Python**:

---

**#  📂 Excel Import**

* User uploads `.xlsm` in Streamlit.
* `pandas.read_excel(..., sheet_name="ENTERED ON")` extracts reservations.
* Data stored in SQLite (`reservations_raw`) with unique `reservation_id`.

---

2. **📧 WIn32 Connection**

   * Use Python’s Win32  to connect to the mailbox.
   * Search last **2 days**:

     mail.search(None,
   * Fetch all emails, filter by subject/body containing `guest name` or `reservation number`.

---

3. **# 📎 Extract PDF Attachments**

   * Download attached PDFs from matching emails.
   * Use `pdfplumber` or `PyPDF2` to extract text.
   * Store raw text into `email_index` table for parsing.

---

4. **🔎 Regex-Based Field Extraction**

   * Apply regex patterns on PDF text:

     * `FULL_NAME`: `r"Name[:\s]+(.+)"`
     * `ARRIVAL`: `r"Arrival[:\s]+(\d{2}/\d{2}/\d{4})"`
     * `DEPARTURE`: `r"Departure[:\s]+(\d{2}/\d{2}/\d{4})"`
     * `NIGHTS`: `r"Nights[:\s]+(\d+)"`
     * `PERSONS`: `r"Persons[:\s]+(\d+)"`
     * `ROOM`: `r"Room[:\s]+(\w+)"`
     * `RATE_CODE`: `r"Rate Code[:\s]+(\w+)"`
     * `COMPANY`: `r"Company[:\s]+(.+)"`
     * `NET TOTAL`: `r"Total[:\s]+([\d,]+\.?\d*)"`
     * `TDF`: `r"TDF[:\s]+([\d,]+\.?\d*)"`

---

5. **📊 Merge with Excel Data**

   * For each Excel reservation → search parsed email data.
   * Matching rule: `guest name + arrival date`.
   * Fill missing Excel fields with PDF-extracted values.
   * If still missing → mark as `N/A`.

---

6. **✅ Auditing Logic**

   * Validation checks:

     * `NIGHTS = Departure - Arrival`.
     * `NET_TOTAL >= TDF`.
     * `PERSONS > 0`.
     * If invalid → mark field as `N/A` + flag record as `FAIL`.
   * Store results in `reservations_enriched` table.

---

7. **🖥️ Streamlit UI**

   * **Step 1:** Upload `.xlsm` → preview reservations.
   * **Step 2:** “Fetch Emails” → connect via IMAP, parse last 2 days.
   * **Step 3:** “Run Audit” → enrich & validate reservations.
   * **Step 4:** Show results in editable data grid.
   * **Step 5:** Export to CSV/XLSX.

---

8. **📋 Error Handling & Logging**

   * If no email match → log `"NO EMAIL FOUND"`.
   * If regex fails → `"FIELD NOT FOUND → N/A"`.
   * Streamlit panel shows processing logs + counts (e.g., `100 reservations → 82 complete, 18 N/A`).

---

9. **🔐 Security & Setup**

   * Store IMAP credentials in `.streamlit/secrets.toml`.
   * Option to configure mail server (`imap.gmail.com`, `outlook.office365.com`, etc.).
   * Ensure PDFs are deleted after parsing (GDPR/data privacy).

---

10. **📤 Final Extraction**

* Export enriched dataset with fields:

  ```
  FULL_NAME, FIRST_NAME, ARRIVAL, DEPARTURE, NIGHTS, PERSONS, ROOM, TDF, NET_TOTAL, RATE_CODE, COMPANY, audit_status
  ```
* Include `source_flags` (Excel, Email, N/A) for transparency.
* Provide audit trail log (CSV) for management review.

**Rate structure**

TDF - Number of nights *20 or 40 for Two Bedroom Apartment (2BA). for 30 days and above it is 30**20 or 40 depeding on room type

Net- Rate with Taxes excluding TDF

Total- Rate with Taxes and TDF

Amount- Rate without Taxes

ADR- Average daily rate Rate without taxes

# OTA Reservations

1. Bookings from T- Booking.com, T- Expedia, T- Agoda.com Company, Brand.com
2. Bookings from *INNLINK2WAY under the  under the INSERT_USER label will be from  email
3. INNLINK2WAY the date format is the csame and has to conversted from mm/dd/yyyy to dd/mm/yyyy
4. for T-Booking.com and T-Brand.com the amount on email will be Mail_TOTAL . For such reservations Mail_NET_TOTAL = Mail_TOTAL-Mail_TDF. amount with TDF. MAIL AMOUNT = Mail_NET_TOTAL/1.225
5. For T- Expedia and Agoda the amount in the email will be Mail_NET_TOTAL. amount without TDF. Mail_TOTAL=Mail_NET_TOTAL+Mail_TDF. MAIL AMOUNT = Mail_NET_TOTAL/1.225. Mail_ADR= MAIL AMOUNT/MAIL_NIGHTS

# Travel Agency Reservations

## Dubai Link (Global Travel Engine)

**Email Format**: Confirmed Booking with Ref. No. [BOOKING_CODE]
**Sender**: suppliers@gte.travel

### Field Mapping:
- **MAIL_FIRST_NAME**: From "Name:" field (e.g., SOHEIL)
- **MAIL_FULL_NAME**: From "Last Name:" field (e.g., RADIOM)
- **MAIL_ARRIVAL**: From "Arrival Date:" (dd/mm/yyyy format)
- **MAIL_DEPARTURE**: From "Departure Date:" (dd/mm/yyyy format)
- **MAIL_NIGHTS**: Calculated from arrival/departure dates
- **MAIL_PERSONS**: From "Adult" count in room description
- **MAIL_ROOM**: From "Rooms:" field (e.g., "1 x Superior Room (King/Twin) - Double (1 Adult)")
- **MAIL_RATE_CODE**: From "Promo code:" field (e.g., TOBBJN{ALL MARKET EX UAE})
- **MAIL_C_T_S**: "Dubai Link" (travel agency name)
- **MAIL_NET_TOTAL**: From "Booking cost price:" field
- **MAIL_TDF**: Calculated as (nights × 20) for regular rooms, (nights × 40) for 2BA rooms
- **MAIL_TOTAL**: MAIL_NET_TOTAL + MAIL_TDF
- **MAIL_AMOUNT**: MAIL_NET_TOTAL ÷ 1.225 (amount without taxes)
- **MAIL_ADR**: MAIL_AMOUNT ÷ MAIL_NIGHTS (average daily rate)

### TDF Calculation Logic:
- Regular rooms: nights × 20 AED
- Two Bedroom Apartments (2BA): nights × 40 AED
- For stays 30+ nights: use 30 × rate (cap at 30 nights)

### Sample Extraction:
```
Name: SOHEIL → MAIL_FIRST_NAME: SOHEIL
Last Name: RADIOM → MAIL_FULL_NAME: RADIOM
Arrival Date: 27/08/2025 → MAIL_ARRIVAL: 27/08/2025
Departure Date: 28/08/2025 → MAIL_DEPARTURE: 28/08/2025
1 x Superior Room (King/Twin) - Double (1 Adult) → MAIL_PERSONS: 1
Booking cost price: 200.00 AED → MAIL_NET_TOTAL: 200.00
Promo code: TOBBJN{ALL MARKET EX UAE} → MAIL_RATE_CODE: TOBBJN{ALL MARKET EX UAE}
``` 

Great question 👌 — let’s clarify exactly  **how SQLite fits into your audit project** .

---

## 🔹 Why Use SQLite?

* **Lightweight DB** (just a `.db` file).
* Works fully  **offline** .
* Lets you keep a  **history of audits** , not just one-off runs.
* Enables **fast filtering/searching** in Streamlit (instead of re-parsing Excel + Outlook every time).
* Provides an **audit trail** (who changed what, when).

---

## 🔹 Where SQLite Fits in the Workflow

1. **🟢 Load Excel → Raw Table**

   * When you upload the `.xlsm` **ENTERED ON** sheet, you write it into SQLite as a table `reservations_raw`.

   ```python
   df.to_sql("reservations_raw", conn, if_exists="replace", index=False)
   ```

   This gives you a persistent store of what came from Excel.

---

2. **🟢 Fetch Emails → Parsed Table**
   * As you extract reservation data from Outlook emails/PDFs, you insert them into `reservations_email`.
   * Schema: use outlook extraction schema

---

3. **🟢 Enrichment → Final Table**
   * Merge `reservations_raw` + `reservations_email`.
   * Fill missing fields (from email or default `N/A`).
   * Apply audit rules.
   * Store results in `reservations_audit`.
   * Schema example: Use Audit Results table  schema

---

4. **🟢 Logs & History**
   * Each run inserts into an `audit_log` table:
     * Timestamp
     * How many reservations loaded
     * How many failed audit
     * Missing fields count

---

5. **🟢 Streamlit Queries SQLite**
   * Instead of working on DataFrames only, Streamlit can:
     * Query reservations by date, status, or company (`SELECT * FROM reservations_audit WHERE audit_status='FAIL'`)
     * Filter by `arrival_date BETWEEN X AND Y`.
     * Show dashboards (missing fields by type).

---

6. **🟢 Export**
   * Final results come from `reservations_audit`.
   * User clicks “Export” → query SQLite → output CSV/XLSX.

---

## 🔹 Example Flow with SQLite

```
Excel (.xlsm) ─▶ reservations_raw
Outlook/PDFs  ─▶ reservations_email
              ─▶ merge + audit ─▶ reservations_audit
Logs & runs   ─▶ audit_log
```

---

## 🔹 Why Not Just Use Pandas?

* Pandas is fine for  **one-off runs** , but SQLite gives you:
  * **Persistence** (results saved even after app closes).
  * **Filtering/Searching** large data much faster.
  * **History** (you can compare today’s vs yesterday’s audit).
  * **Integration** (easier to plug into BI dashboards later).

---

👉 Do you want me to show you a **sample SQLite schema + code snippets** for `reservations_raw`, `reservations_email`, and `reservations_audit` so you see how the tables will look?

---

👉 Do you want me to **draft a Streamlit starter code** (+INSTALL INSTY  REGREGRE PDF extraction + regex parsing + DB load) so you can test the workflow end-to-end?
