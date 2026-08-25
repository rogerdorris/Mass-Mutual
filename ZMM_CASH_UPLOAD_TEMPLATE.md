# ZMM_CASH_UPLOAD_TEMPLATE – Excel Upload Template Layout

## Purpose
Standardized Excel template for mass upload of customer payments to SAP S/4HANA.
Used by report `ZFIAR_MASS_CASH_POSTING` to perform bulk F-28 cash posting and A/R clearing.

---

## File Format
- File type: `.xlsx` (Excel 2007+)
- **Row 1**: Column headers (do not modify)
- **Row 2 onward**: Data rows (one row per payment/clearing entry)
- Maximum rows: 99,999

---

## Column Layout

| Col # | Header Label          | SAP Field         | Format          | Required | Example          | Notes |
|-------|-----------------------|-------------------|-----------------|----------|------------------|-------|
| 1     | Company Code          | BKPF-BUKRS        | Text (4 chars)  | Yes      | `2920` or `1711` | Must match valid company code |
| 2     | Customer ID           | RF05A-AGKON       | Text (10 chars) | Yes      | `0000100050`     | Leading zeros required |
| 3     | Invoice Number(s)     | RFOPS_DK-BELNR    | Text (255 chars)| Yes      | `0093818812` or `0093818812;0093818813` | One or more invoice numbers separated by `;`. Each must be open in BSID. |
| 4     | Payment Amount        | BSEG-WRBTR        | Decimal         | Yes      | `100.00`         | Total payment; must be ≤ sum of all listed invoice balances |
| 5     | Payment Date          | BKPF-BLDAT        | YYYYMMDD        | Yes      | `20260825`       | Document date |
| 6     | Posting Date          | BKPF-BUDAT        | YYYYMMDD        | Yes      | `20260825`       | Accounting posting date |
| 7     | Value Date            | BSEG-VALUT        | YYYYMMDD        | Yes      | `20260825`       | Bank value date |
| 8     | Currency              | BKPF-WAERS        | Text (3 chars)  | Yes      | `USD` or `CAD`   | ISO currency code |
| 9     | House Bank            | BSEG-HBKID        | Text (5 chars)  | Yes      | `BOA11`          | House bank ID configured in FICA |
| 10    | House Bank Account ID | BSEG-HKTID        | Text (5 chars)  | Yes      | `DEP01`          | Bank account sub-ID |
| 11    | G/L Account           | RF05A-KONTO       | Text (10 chars) | Yes      | `0011001200`     | Payment processor clearing G/L (Stripe Clearing) |
| 12    | Text / Reference      | BSEG-SGTXT        | Text (50 chars) | No       | `Stripe payout August 2026` | Line item text |
| 13    | Transaction ID        | BKPF-XBLNR        | Text (50 chars) | No       | `STR-20260825-001` | External reference for duplicate prevention |

---

## Validation Rules Applied by the Program

| Rule | Description | Error Message |
|------|-------------|---------------|
| Required fields | Columns 1–4, 8–11 must not be blank (after selection-screen overrides are applied) | `Missing required field(s): Company Code / Customer / Invoice(s) / Amount / Currency / House Bank / House Bank Acct / G/L Account` |
| Customer existence | Customer ID must exist in table KNA1 | `Customer ID Not Found in KNA1` |
| Open invoice (each) | Every invoice number must be open in BSID for that customer/company | `Invoice XXXXXXXXXX Not Found in open items (BSID)` |
| Already cleared (each) | Invoice must not be in BSAD | `Invoice XXXXXXXXXX Already Cleared (exists in BSAD)` |
| Amount over-payment | Payment amount must not exceed sum of all listed invoice balances | `Payment exceeds total open balance` |
| Duplicate transaction | Transaction ID must not already exist in BKPF-XBLNR | `Duplicate Entry – Transaction ID already posted` |
| Partial payment | Payment < total invoices: residual open item auto-created on last invoice | Warning: `partial payment, residual X will be created` |

---

## G/L Posting Logic

```
Debit:   Payment Processor Clearing G/L  (Col 11, e.g., 11001200 – Stripe Clearing)
Credit:  Customer A/R Account            (Col 2,  clears open invoice from Col 3)
```

Document Type: **DZ** (Customer Payment)

---

## Sample Data Row

| 1      | 2            | 3          | 4      | 5        | 6        | 7        | 8   | 9     | 10    | 11         | 12                   | 13                |
|--------|--------------|------------|--------|----------|----------|----------|-----|-------|-------|------------|----------------------|-------------------|
| `2920` | `0000100050` | `0093818812;0093818813` | `200.00` | `20260825` | `20260825` | `20260825` | `USD` | `BOA11` | `DEP01` | `0011001200` | `Stripe Aug payout` | `STR-20260825-001` |

---

## Run-Time Field Overrides

The selection screen contains a **Field Overrides** block. Any parameter left blank means "use the value from the file." A non-blank entry replaces the corresponding column for **every row** in the batch before validation or posting.

| Screen Parameter | Type    | File Column Overridden    | SAP Field        | Notes |
|------------------|---------|---------------------------|------------------|-------|
| `P_PDATE`        | `DATS`  | Col 6 – Posting Date      | `BKPF-BUDAT`     | Force today's date when running a file prepared in advance |
| `P_DDATE`        | `DATS`  | Col 5 – Payment Date      | `BKPF-BLDAT`     | Override document/payment date for late-upload scenarios |
| `P_VDATE`        | `VALUT` | Col 7 – Value Date        | `BSEG-VALUT`     | Bank value date; usually uniform across a batch |
| `P_CURR`         | `WAERS` | Col 8 – Currency          | `BKPF-WAERS`     | Scope a multi-currency file to a single currency |
| `P_HBKID`        | `HBKID` | Col 9 – House Bank        | `BSEG-HBKID`     | Switch house bank without editing the file |
| `P_HKTID`        | `HKTID` | Col 10 – House Bank Acct  | `BSEG-HKTID`     | Switch bank account sub-ID |
| `P_GLACC`        | `SAKNR` | Col 11 – G/L Account      | `RF05A-KONTO`    | Switch clearing G/L (e.g., Stripe vs. PayPal) |
| `P_DTYPE`        | `BLART` | *(hardcoded `DZ`)*        | `BKPF-BLART`     | Default `DZ`; change only when an alternate document type is required |

---

## Notes

- **Test Mode**: Select the *Test Run* checkbox in the program selection screen to validate all rows without creating SAP documents.
- **Partial Payments**: When payment amount is less than the invoice balance, the difference is automatically posted as a residual open item (`DF05B-PSDIF` logic).
- **Post-Upload Step**: After mass upload, Electronic Bank Statements (EBS) clear the payment processor clearing G/L against the primary Cash G/L account (separate process).
- **Transport**: Program objects must be assigned to a transport request and promoted Dev → QA → PROD.
