# 🚀 **Reservation Audit Pipeline (Streamlit + Python)**

### **Step 1: Tech Stack**

* **Data Extraction**
  * Or **Outlook COM API (pywin32)** if running on Windows desktop with Outlook installed.
  * `pdfplumber` → Extract text from PDF attachments.
  * `beautifulsoup4` → Clean HTML emails.
* **Processing**
  * `regex` + `dateparser` → Extract names, dates, rate codes.
  * `spaCy` (NER) → Extract person names, organizations.
* **Database**
  * `sqlite3` → Store structured reservations.
* **App**
  * `streamlit` → User interface to upload emails, run audits, and view results.


### **Step 2: Pipeline Flow**

1. **Load Email Data****Normalize Attachments**
2. * If PDF → use `pdfplumber` → text.
   * If HTML → `BeautifulSoup` → text.
   * Else → plain text.
3. **Extract Fields**
   * Regex + `dateparser`:

     * Dates, numbers, nights, persons.
   * `spaCy`:

     * Names, companies.
   * Combine results into dictionary:
   * 
   * **Audit Fields**

     * If field missing → mark as `N/A`.
   * **Save to SQLite**

     * Table `reservations_audit` with all fields.
     * Easy queries + history retention.
   * **Streamlit UI**

     * Upload emails.
     * Click "Run Audit".
     * Display results in interactive table.
     * Download CSV of audited reservations.
4. 

# **Summary of Pipeline**

1. **Input** : Emails (text, HTML, PDF).
2. **Extraction** : `pdfplumber` + regex + spaCy.
3. **Audit** : Mark missing as `N/A`.
4. **Storage** : SQLite DB.
5. **UI** : Streamlit to upload, run, and export.
