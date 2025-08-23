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

---

👉 Do you want me to **draft a Streamlit starter code** (+INSTALL INSTY  REGREGRE PDF extraction + regex parsing + DB load) so you can test the workflow end-to-end?
