# README — NER Training for Reservation Email Extraction

1. **Overview — Purpose & Scope**
   Build a robust NER (Named Entity Recognition) model that extracts reservation fields from Outlook emails (mixed formats: plain text, HTML, PDFs). The model will learn your agent-specific formats and output fields like `MAIL_FIRST_NAME`, `MAIL_FULL_NAME`, `MAIL_ARRIVAL`, `MAIL_DEPARTURE`, `MAIL_NIGHTS`, `MAIL_PERSONS`, `MAIL_ROOM`, `MAIL_RATE_CODE`, `MAIL_C_T_S`, `MAIL_NET_TOTAL`, `MAIL_TOTAL`, `MAIL_TDF`, `MAIL_ADR`, `MAIL_AMOUNT`. Use your existing Python parser to bootstrap labels, correct them, fine-tune a transformer NER on Colab GPU, then run inference locally (Streamlit + SQLite).

2. **Labels — Canonical Field List**
   Use these exact labels everywhere (training, DB, pipeline). Example canonical set you provided:
   `MAIL_FIRST_NAME`, `MAIL_FULL_NAME`, `MAIL_ARRIVAL`, `MAIL_DEPARTURE`, `MAIL_NIGHTS`, `MAIL_PERSONS`, `MAIL_ROOM`, `MAIL_RATE_CODE`, `MAIL_C_T_S`, `MAIL_NET_TOTAL`, `MAIL_TOTAL`, `MAIL_TDF`, `MAIL_ADR`, `MAIL_AMOUNT`.
   Map any parser variants to these canonical names before converting to token-level labels.

3. **Data Collection & Weak Labeling (Bootstrap)**

   * Run your existing Python parser on a large batch of historical emails to produce JSON records (one per email) with the canonical fields. These are **weak labels**.
   * Save each record with `email_id`, `raw_text`, and the parser’s field values. This gives scale quickly and reduces manual work.

4. **Convert to Token-Level BIO Format**

   * Convert each labeled email to tokenized sequences with BIO tags. Each token receives one of: `B-MAIL_X`, `I-MAIL_X`, or `O`.
   * Minimal conversion example (pseudo-code):

     ```python
     tokens = text.split()  # or use a tokenizer
     labels = ["O"] * len(tokens)
     for field, value in parsed_fields.items():
         if value and value != "N/A":
             start_idx, end_idx = find_token_span(tokens, value)  # robust matching
             labels[start_idx] = f"B-{field}"
             for i in range(start_idx+1, end_idx+1):
                 labels[i] = f"I-{field}"
     save_conll(tokens, labels)
     ```
   * Save train/test/validation splits (e.g., 80/10/10). Keep chronological split if distribution changes over time.

5. **Human-in-the-Loop Annotation & Correction**

   * Load the weak-labeled data into an annotation tool (doccano or Label Studio).
   * Annotators correct mistakes and confirm boundaries (focus on edge cases: date formats, amounts with currency symbols, multi-word names).
   * Export corrected dataset in CoNLL/JSON for training.

6. **Model Selection & Colab Training Setup**

   * Recommended base models: `distilbert-base-cased` (fast) or `bert-base-cased` (accurate).
   * Use Google Colab (Runtime → GPU). Install: `pip install transformers datasets seqeval`.
   * Steps: tokenizer → tokenize+align labels → `AutoModelForTokenClassification` → `Trainer` with `compute_metrics` using `seqeval`. Typical hyperparams: lr=2e-5, batch=8–16, epochs=3–5. Save model artifacts after training.

7. **Tokenization & Label Alignment (important detail)**

   * Use `is_split_into_words=True` when tokenizing pre-split tokens or align word IDs to labels. For subword tokens, set label = `-100` for non-first subword to ignore in loss. Example code pattern is provided in HuggingFace docs—use that alignment function exactly to avoid mis-labeled training.

8. **Evaluation, Thresholding & Error Analysis**

   * Evaluate per-entity precision/recall/F1 (use `seqeval`). Track per-label F1 (dates vs amounts vs names).
   * Use confidence scores at inference; set thresholds (e.g., accept predictions with score ≥ 0.8). Low-confidence fields go to manual review in Streamlit.
   * Maintain an error log: common false positives/negatives (e.g., currency symbols, ambiguous date formats). Use these to augment training data.

9. **Deployment & Integration**

   * Export trained model artifacts from Colab (`trainer.save_model("reservation-ner")`) and download ZIP. Locally, load with HuggingFace pipeline:

     ```python
     from transformers import pipeline
     ner = pipeline("ner", model="reservation-ner", tokenizer="reservation-ner", aggregation_strategy="simple")
     entities = ner(email_text)
     ```
   * Post-process entities into canonical fields (normalize dates with `dateparser`, amounts with `decimal` after removing currency symbols). Merge into SQLite reservations table and follow audit rules (fill missing → `N/A`, compute `NIGHTS`, validate `MAIL_NET_TOTAL >= MAIL_TDF`, etc.). Surface low-confidence / conflicting records in Streamlit data grid.

10. **Lifecycle: Retraining, Versioning & Monitoring**

* **Retrain cadence:** monthly or when >X new distinctive email formats appear. Use the human corrections saved from Streamlit as additional gold data.
* **Versioning:** tag models with date and dataset used (e.g., `reservation-ner_v2025-09-04`). Keep model and dataset hashes.
* **Monitoring:** track extraction accuracy over time (per-field F1), percent of fields flagged `N/A`, and manual correction rate. Use that to trigger re-annotation + retraining.
* **Fallbacks:** keep original regex parser as a fallback validator (compare parser vs model outputs and flag conflicts).

---

## Appendix — Quick Practical Tips

* Normalization: convert dates to ISO (`YYYY-MM-DD`), amounts to float AED/USD by removing commas and currency symbols.
* Ambiguities: if `MAIL_ARRIVAL` > `MAIL_DEPARTURE` → swap or flag.
* CPU-only inference is fast; heavy training done on Colab GPU.
* Keep sample sizes balanced: collect varied agent samples (PDF-heavy, HTML-heavy, plain-text) so the model generalizes.

---

If you want, I can now:

* produce a **small Python script** to convert your parser’s JSON output into CoNLL/BIO format, or
* create a **Colab-ready notebook** (cells for data upload, tokenization, training, saving) prefilled with your canonical labels. Which do you prefer?
