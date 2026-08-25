*&---------------------------------------------------------------------*
*& Report  ZFIAR_MASS_CASH_POSTING
*& Description: Mass Upload of Customer Payment and Clearing
*&              Emulates F-28 via POSTING_INTERFACE_CLEARING /
*&              POSTING_INTERFACE_DOCUMENT for bulk A/R cash posting.
*&              Supports multiple invoice numbers per row (semicolon-
*&              separated) in the Invoice Numbers column.
*&
*& Upload Modes
*&   FRONTEND – F4 browses the Windows client filesystem via
*&              cl_gui_frontend_services; file is transferred to the
*&              application server using GUI_UPLOAD before parsing.
*&   SERVER   – F4 browses the SAP application server (AL11 style) via
*&              F4_FILENAME; file already resides on the server and is
*&              read directly with CL_FDT_XL_SPREADSHEET.
*&
*& Text symbols to maintain in SE32 / SE38:
*&   TEXT-001 = 'Run Parameters'
*&   TEXT-002 = 'Overpayment Handling'
*&   TEXT-003 = 'Allow overpayment (warning – excess credited on-account)'
*&   TEXT-004 = 'Reject overpayment (error – row will not be posted)'
*&   TEXT-005 = 'File Handling'
*&   TEXT-006 = 'Frontend (Windows directory – GUI upload)'
*&   TEXT-007 = 'Server path (AL11 folder – application server)'
*&   TEXT-008 = 'Output Options'
*&   TEXT-009 = 'Job log (write summary to job/spool log)'
*&   TEXT-010 = 'Spool / ALV list (display on screen)'
*&   TEXT-011 = 'Email (send results report via BCS)'
*&   TEXT-012 = 'Recipient e-mail address'
*&   TEXT-015 = 'Recipient name'
*&   TEXT-013 = 'Windows file path'
*&   TEXT-014 = 'Server file path (AL11)'
*&   TEXT-016 = 'Field Overrides (blank = use file value)'
*&   TEXT-017 = 'Override posting date (BKPF-BUDAT)'
*&   TEXT-018 = 'Override payment/document date (BKPF-BLDAT)'
*&   TEXT-019 = 'Override value date (BSEG-VALUT)'
*&   TEXT-020 = 'Override currency (BKPF-WAERS)'
*&   TEXT-021 = 'Override house bank'
*&   TEXT-022 = 'Override house bank account ID'
*&   TEXT-023 = 'Override G/L clearing account'
*&   TEXT-024 = 'Override document type'
*&   TEXT-025 = 'Test run (simulation – no documents posted)'
*&   TEXT-026 = 'Live run (documents will be posted)'
*&---------------------------------------------------------------------*
REPORT zfiar_mass_cash_posting
  NO STANDARD PAGE HEADING
  LINE-SIZE 255
  MESSAGE-ID 00.

*----------------------------------------------------------------------*
* Type Definitions
*----------------------------------------------------------------------*
TYPES:
  " Upload file row structure
  BEGIN OF ty_upload,
    company_code   TYPE bukrs,       " BKPF-BUKRS  e.g. 2920 / 1711
    customer_id    TYPE kunnr,       " RF05A-AGKON
    invoice_refs   TYPE char255,     " RFOPS_DK-BELNR  – one or more invoice
                                     "   numbers separated by ';'
                                     "   e.g. '0093818812;0093818813;0093818814'
    payment_amount TYPE wrbtr,       " BSEG-WRBTR  total payment amount
    invoice_amount TYPE wrbtr,       " Sum of open item amounts from BSID
    payment_date   TYPE dats,        " BKPF-BLDAT
    posting_date   TYPE dats,        " BKPF-BUDAT
    value_date     TYPE valut,       " BSEG-VALUT
    currency       TYPE waers,       " BKPF-WAERS  e.g. USD / CAD
    house_bank     TYPE hbkid,       " BSEG-HBKID  e.g. BOA11
    house_bank_id  TYPE hktid,       " BSEG-HKTID  e.g. DEP01
    gl_account     TYPE saknr,       " RF05A-KONTO e.g. 11001200
    text           TYPE sgtxt,       " BSEG-SGTXT
    transaction_id TYPE char50,      " External transaction ID (dedup key)
  END OF ty_upload,

  " Individual validated invoice line derived from one upload row
  BEGIN OF ty_invoice,
    row_num        TYPE i,
    invoice_ref    TYPE belnr_d,
    open_amount    TYPE wrbtr,       " Amount from BSID
    fiscal_year    TYPE gjahr,
  END OF ty_invoice,

  " Execution log line structure
  BEGIN OF ty_log,
    row_num        TYPE i,
    company_code   TYPE bukrs,
    customer_id    TYPE kunnr,
    invoice_refs   TYPE char255,     " All invoice numbers for this row
    payment_amount TYPE wrbtr,
    residual_amt   TYPE wrbtr,
    sap_doc_num    TYPE belnr_d,
    traffic_light  TYPE c,           " ALV traffic light: 1=red 2=yellow 3=green
    icon           TYPE icon_d,      " ALV icon (ICON_LED_RED / YELLOW / GREEN)
    status         TYPE char10,      " SUCCESS / WARNING / ERROR / SIM-OK
    message        TYPE char200,
  END OF ty_log.

*----------------------------------------------------------------------*
* Internal Tables & Work Areas
*----------------------------------------------------------------------*
DATA:
  gt_upload   TYPE STANDARD TABLE OF ty_upload,
  gs_upload   TYPE ty_upload,
  gt_log      TYPE STANDARD TABLE OF ty_log,
  gs_log      TYPE ty_log,
  gt_raw      TYPE STANDARD TABLE OF alsmex_tabline,
  gs_raw      TYPE alsmex_tabline,
  gt_invoices TYPE STANDARD TABLE OF ty_invoice,  " validated invoices per row
  gs_invoice  TYPE ty_invoice.

DATA:
  gv_file_path  TYPE string,
  gv_total_ok   TYPE i,
  gv_total_err  TYPE i,
  gv_total_amt  TYPE wrbtr.

*----------------------------------------------------------------------*
* ALV References
*----------------------------------------------------------------------*
DATA:
  go_alv      TYPE REF TO cl_salv_table,
  go_columns  TYPE REF TO cl_salv_columns_table,
  go_column   TYPE REF TO cl_salv_column_table,
  go_display  TYPE REF TO cl_salv_display_settings,
  go_funcs    TYPE REF TO cl_salv_functions_list.

*----------------------------------------------------------------------*
* Selection Screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b0 WITH FRAME TITLE TEXT-001.
  PARAMETERS:
    p_bukrs  TYPE bukrs DEFAULT '2920'.                " Company code
  SELECTION-SCREEN SKIP 1.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_tst    RADIOBUTTON GROUP tmod DEFAULT 'X'.      " Test run (simulation)
  SELECTION-SCREEN COMMENT 3(55) TEXT-025 FOR FIELD p_tst.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_live   RADIOBUTTON GROUP tmod.                  " Live run (documents posted)
  SELECTION-SCREEN COMMENT 3(55) TEXT-026 FOR FIELD p_live.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN SKIP 1.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_ovwrn  RADIOBUTTON GROUP ovpy DEFAULT 'X'.   " Allow – warn and credit customer
  SELECTION-SCREEN COMMENT 3(72) TEXT-003 FOR FIELD p_ovwrn.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_overr  RADIOBUTTON GROUP ovpy.               " Reject – treat as error
  SELECTION-SCREEN COMMENT 3(72) TEXT-004 FOR FIELD p_overr.
  SELECTION-SCREEN END OF LINE.
SELECTION-SCREEN END OF BLOCK b0.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-016.
  PARAMETERS:
    p_pdate  TYPE dats,                                " Override posting date
    p_ddate  TYPE dats,                                " Override payment / document date
    p_vdate  TYPE valut,                               " Override value date
    p_curr   TYPE waers,                               " Override currency
    p_hbkid  TYPE hbkid,                               " Override house bank
    p_hktid  TYPE hktid,                               " Override house bank account ID
    p_glacc  TYPE saknr,                               " Override G/L clearing account
    p_dtype  TYPE blart DEFAULT 'DZ'.                  " Override document type
SELECTION-SCREEN END OF BLOCK b3.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-005.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_front  RADIOBUTTON GROUP src DEFAULT 'X'.    " Frontend Windows directory
  SELECTION-SCREEN COMMENT 3(55) TEXT-006 FOR FIELD p_front.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_srvr   RADIOBUTTON GROUP src.                " AL11 application-server folder
  SELECTION-SCREEN COMMENT 3(55) TEXT-007 FOR FIELD p_srvr.
  SELECTION-SCREEN END OF LINE.
  PARAMETERS:
    p_fefile TYPE string,                           " Frontend: Windows file path
    p_srvfl  TYPE string.                           " Server:   AL11 file path
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-008.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_ojob   RADIOBUTTON GROUP out DEFAULT 'X'.    " Output: job log only
  SELECTION-SCREEN COMMENT 3(55) TEXT-009 FOR FIELD p_ojob.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_ospl   RADIOBUTTON GROUP out.                " Output: spool / ALV list
  SELECTION-SCREEN COMMENT 3(55) TEXT-010 FOR FIELD p_ospl.
  SELECTION-SCREEN END OF LINE.
  SELECTION-SCREEN BEGIN OF LINE.
  PARAMETERS:
    p_oeml   RADIOBUTTON GROUP out.                " Output: email via BCS
  SELECTION-SCREEN COMMENT 3(55) TEXT-011 FOR FIELD p_oeml.
  SELECTION-SCREEN END OF LINE.
  PARAMETERS:
    p_email  TYPE ad_smtpadr LOWER CASE.           " Recipient address (email mode)
  PARAMETERS:
    p_ename  TYPE ad_name1.                        " Recipient display name (email mode)
SELECTION-SCREEN END OF BLOCK b2.

*----------------------------------------------------------------------*
* Dynamic screen:
*   – Show p_fefile (Windows path) only when Frontend radio is active
*   – Show p_srvfl  (AL11 path)    only when Server radio is active
*   – Grey-out p_email / p_ename unless Email output radio is chosen
*----------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.
    CASE screen-name.
      WHEN 'P_FEFILE'.
        " Visible and editable only in frontend mode
        IF p_front = 'X'.
          screen-active = '1'.
          screen-input  = '1'.
        ELSE.
          screen-active = '0'.
        ENDIF.
        MODIFY SCREEN.
      WHEN 'P_SRVFL'.
        " Visible and editable only in server mode
        IF p_srvr = 'X'.
          screen-active = '1'.
          screen-input  = '1'.
        ELSE.
          screen-active = '0'.
        ENDIF.
        MODIFY SCREEN.
      WHEN 'P_EMAIL'.
        IF p_oeml = 'X'.
          screen-input       = '1'.
          screen-intensified = '0'.
        ELSE.
          screen-input       = '0'.
          screen-intensified = '1'.
        ENDIF.
        MODIFY SCREEN.
      WHEN 'P_ENAME'.
        IF p_oeml = 'X'.
          screen-input       = '1'.
          screen-intensified = '0'.
        ELSE.
          screen-input       = '0'.
          screen-intensified = '1'.
        ENDIF.
        MODIFY SCREEN.
    ENDCASE.
  ENDLOOP.

*----------------------------------------------------------------------*
* Validation: enforce that the visible path field is filled, and that
* an email address is provided when email output is selected
*----------------------------------------------------------------------*
AT SELECTION-SCREEN.
  IF p_front = 'X' AND p_fefile IS INITIAL.
    MESSAGE e001(00) WITH 'Enter a Windows file path.'
      DISPLAY LIKE 'E'.
  ENDIF.
  IF p_srvr = 'X' AND p_srvfl IS INITIAL.
    MESSAGE e001(00) WITH 'Enter a server (AL11) file path.'
      DISPLAY LIKE 'E'.
  ENDIF.
  IF p_oeml = 'X' AND p_email IS INITIAL.
    MESSAGE e001(00) WITH 'Enter a recipient e-mail address for email output.'
      DISPLAY LIKE 'E'.
  ENDIF.
  IF p_oeml = 'X' AND p_ename IS INITIAL.
    MESSAGE e001(00) WITH 'Enter a recipient name for email output.'
      DISPLAY LIKE 'E'.
  ENDIF.

*----------------------------------------------------------------------*
* F4 help: each path field has its own dedicated value-request handler
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fefile.
  PERFORM f_browse_file CHANGING p_fefile.        " Windows client directory scan

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_srvfl.
  PERFORM f_browse_al11 CHANGING p_srvfl.         " AL11 application-server folder

*----------------------------------------------------------------------*
* Main Program
*----------------------------------------------------------------------*
START-OF-SELECTION.

  " Authority check
  PERFORM f_authority_check.

  " Upload and parse the Excel file (pass the active path field)
  IF p_front = 'X'.
    PERFORM f_upload_file USING p_fefile.
  ELSE.
    PERFORM f_upload_file USING p_srvfl.
  ENDIF.

  " Validate all rows before posting
  PERFORM f_validate_data USING p_ovwrn.

  " Post or simulate
  PERFORM f_post_payments USING p_tst p_ovwrn.

  " Output: branch based on selected output destination
  IF p_ojob = 'X'.
    " Job log / spool only – write summary text, no ALV pop-up
    PERFORM f_write_job_log.
  ELSEIF p_ospl = 'X'.
    " Spool / ALV list – full screen display
    PERFORM f_display_alv.
  ELSE.
    " Email via BCS – send results report (requires p_email)
    PERFORM f_send_email.
  ENDIF.

*----------------------------------------------------------------------*
* Form: Browse File – Frontend (F4 scans Windows client directories)
*----------------------------------------------------------------------*
FORM f_browse_file CHANGING cv_file TYPE string.
  DATA: lv_rc      TYPE i,
        lt_filetab TYPE filetable,
        ls_filetab LIKE LINE OF lt_filetab.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Select Cash Upload File (Windows)'
      default_extension       = 'xlsx'
      file_filter             = '*.xlsx;*.xls'
    CHANGING
      file_table              = lt_filetab
      rc                      = lv_rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      OTHERS                  = 2.

  IF sy-subrc = 0 AND lv_rc = 1.
    READ TABLE lt_filetab INTO ls_filetab INDEX 1.
    cv_file = ls_filetab-filename.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Browse AL11 – Server (F4 browses application-server folders)
*       Uses F4_FILENAME which renders the standard AL11-style file
*       picker scoped to the SAP application server filesystem.
*----------------------------------------------------------------------*
FORM f_browse_al11 CHANGING cv_file TYPE string.
  DATA: lv_path TYPE string.

  lv_path = cv_file.

  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = syst-repid
      dynpro_number = syst-dynnr
      field_name    = 'P_FILE'
    IMPORTING
      file_name     = lv_path.

  IF lv_path IS NOT INITIAL.
    cv_file = lv_path.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Authority Check
*----------------------------------------------------------------------*
FORM f_authority_check.
  " Check customer display authorization
  AUTHORITY-CHECK OBJECT 'F_KNA1_BUK'
    ID 'BUKRS' FIELD p_bukrs
    ID 'ACTVT' FIELD '03'.
  IF sy-subrc <> 0.
    MESSAGE e001(00) WITH 'No authorization for customer master (F_KNA1_BUK)'.
  ENDIF.

  " Check FI document posting authorization
  AUTHORITY-CHECK OBJECT 'F_BKPF_BUK'
    ID 'BUKRS' FIELD p_bukrs
    ID 'ACTVT' FIELD '01'.
  IF sy-subrc <> 0.
    MESSAGE e001(00) WITH 'No authorization for FI document posting (F_BKPF_BUK)'.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Upload File
*   FRONTEND mode: transfer file from Windows client to app server
*     memory using GUI_UPLOAD, then parse with KCD_EXCEL_OLE_TO_INTERNAL_TABLE.
*   SERVER mode: file already on app server; read as xstring with
*     OPEN DATASET and parse with CL_FDT_XL_SPREADSHEET (no OLE/GUI).
*----------------------------------------------------------------------*
FORM f_upload_file USING iv_file TYPE string.
  DATA: lv_filename TYPE string.

  lv_filename = iv_file.

  IF p_front = 'X'.
    " ---------------------------------------------------------------
    " Frontend path: use OLE-based Excel reader (SAP GUI required)
    " ---------------------------------------------------------------
    CALL FUNCTION 'KCD_EXCEL_OLE_TO_INTERNAL_TABLE'
      EXPORTING
        filename                = lv_filename
        i_begin_col             = 1
        i_begin_row             = 2    " Row 1 = header
        i_end_col               = 14
        i_end_row               = 99999
      TABLES
        intern                  = gt_raw
      EXCEPTIONS
        inconsistent_parameters = 1
        upload_ole              = 2
        OTHERS                  = 3.

    IF sy-subrc <> 0.
      MESSAGE e001(00) WITH 'Error reading Excel file (frontend). Check path and format.'.
      STOP.
    ENDIF.

  ELSE.
    " ---------------------------------------------------------------
    " Server path: read raw xstring from application server file,
    " then parse with CL_FDT_XL_SPREADSHEET (no GUI/OLE dependency).
    " ---------------------------------------------------------------
    DATA: lv_xstr    TYPE xstring,
          lv_buffer  TYPE xstring,
          lo_xl      TYPE REF TO cl_fdt_xl_spreadsheet,
          lt_sheets  TYPE if_fdt_doc_spreadsheet=>t_sheetnames,
          lv_sheet   TYPE string,
          lt_xl_tab  TYPE if_fdt_doc_spreadsheet=>t_data,
          ls_xl_line TYPE if_fdt_doc_spreadsheet=>s_data.

    " Read server file into xstring
    OPEN DATASET lv_filename FOR INPUT IN BINARY MODE.
    IF sy-subrc <> 0.
      MESSAGE e001(00) WITH 'Cannot open server file – check AL11 path and permissions.'.
      STOP.
    ENDIF.
    DO.
      READ DATASET lv_filename INTO lv_buffer.
      IF sy-subrc <> 0. EXIT. ENDIF.
      CONCATENATE lv_xstr lv_buffer INTO lv_xstr IN BYTE MODE.
    ENDDO.
    CLOSE DATASET lv_filename.

    IF lv_xstr IS INITIAL.
      MESSAGE e001(00) WITH 'Server file is empty – check AL11 path.'.
      STOP.
    ENDIF.

    " Parse XLSX on the application server (no OLE)
    TRY.
      CREATE OBJECT lo_xl
        EXPORTING
          iv_data              = lv_xstr
          iv_xlsx              = abap_true.

      lt_sheets = lo_xl->get_sheet_names( ).
      IF lt_sheets IS INITIAL.
        MESSAGE e001(00) WITH 'No worksheets found in server Excel file.'.
        STOP.
      ENDIF.

      " Use first sheet
      READ TABLE lt_sheets INTO lv_sheet INDEX 1.

      lo_xl->if_fdt_doc_spreadsheet~get_sheet_content(
        EXPORTING
          iv_sheet_name = lv_sheet
        IMPORTING
          et_data       = lt_xl_tab ).

    CATCH cx_fdt_xl_spreadsheet.
      MESSAGE e001(00) WITH 'Error parsing server Excel file (CL_FDT_XL_SPREADSHEET).'.
      STOP.
    ENDTRY.

    " Map CL_FDT_XL_SPREADSHEET output to alsmex_tabline format (gt_raw)
    " Skip header row (row_index = 1)
    LOOP AT lt_xl_tab INTO ls_xl_line.
      IF ls_xl_line-row = 1. CONTINUE. ENDIF.   " header

      CLEAR gs_raw.
      gs_raw-row   = ls_xl_line-row.
      gs_raw-col   = ls_xl_line-col.
      gs_raw-value = ls_xl_line-value.
      APPEND gs_raw TO gt_raw.
    ENDLOOP.
  ENDIF.

  IF gt_raw IS INITIAL.
    MESSAGE e001(00) WITH 'No data found in upload file.'.
    STOP.
  ENDIF.

  " Map raw cells to typed upload structure (common for both modes)
  DATA: lv_row_prev TYPE i VALUE 0,
        lv_col      TYPE i.

  LOOP AT gt_raw INTO gs_raw.
    lv_col = gs_raw-col.

    " New row detected – append previous work area
    IF gs_raw-row <> lv_row_prev AND lv_row_prev > 0.
      IF p_bukrs IS INITIAL OR gs_upload-company_code = p_bukrs.
        APPEND gs_upload TO gt_upload.
      ENDIF.
      CLEAR gs_upload.
    ENDIF.

    CASE lv_col.
      WHEN 1.  gs_upload-company_code   = gs_raw-value.
      WHEN 2.  gs_upload-customer_id    = gs_raw-value.
      WHEN 3.  gs_upload-invoice_refs   = gs_raw-value.   " semicolon-separated
      WHEN 4.  gs_upload-payment_amount = gs_raw-value.
      WHEN 5.  gs_upload-payment_date   = gs_raw-value.
      WHEN 6.  gs_upload-posting_date   = gs_raw-value.
      WHEN 7.  gs_upload-value_date     = gs_raw-value.
      WHEN 8.  gs_upload-currency       = gs_raw-value.
      WHEN 9.  gs_upload-house_bank     = gs_raw-value.
      WHEN 10. gs_upload-house_bank_id  = gs_raw-value.
      WHEN 11. gs_upload-gl_account     = gs_raw-value.
      WHEN 12. gs_upload-text           = gs_raw-value.
      WHEN 13. gs_upload-transaction_id = gs_raw-value.
    ENDCASE.

    lv_row_prev = gs_raw-row.
  ENDLOOP.

  " Append last row
  IF gs_upload IS NOT INITIAL.
    IF p_bukrs IS INITIAL OR gs_upload-company_code = p_bukrs.
      APPEND gs_upload TO gt_upload.
    ENDIF.
  ENDIF.

  IF gt_upload IS INITIAL.
    MESSAGE e001(00) WITH 'No matching rows found for company code ' && p_bukrs.
    STOP.
  ENDIF.

  " Apply selection-screen overrides to every loaded row (blank = keep file value)
  IF p_pdate IS NOT INITIAL OR p_ddate IS NOT INITIAL OR p_vdate IS NOT INITIAL
  OR p_curr  IS NOT INITIAL OR p_hbkid IS NOT INITIAL OR p_hktid IS NOT INITIAL
  OR p_glacc IS NOT INITIAL.
    LOOP AT gt_upload INTO gs_upload.
      IF p_pdate IS NOT INITIAL. gs_upload-posting_date  = p_pdate. ENDIF.
      IF p_ddate IS NOT INITIAL. gs_upload-payment_date  = p_ddate. ENDIF.
      IF p_vdate IS NOT INITIAL. gs_upload-value_date    = p_vdate. ENDIF.
      IF p_curr  IS NOT INITIAL. gs_upload-currency      = p_curr.  ENDIF.
      IF p_hbkid IS NOT INITIAL. gs_upload-house_bank    = p_hbkid. ENDIF.
      IF p_hktid IS NOT INITIAL. gs_upload-house_bank_id = p_hktid. ENDIF.
      IF p_glacc IS NOT INITIAL. gs_upload-gl_account    = p_glacc. ENDIF.
      MODIFY gt_upload FROM gs_upload.
    ENDLOOP.
  ENDIF.

  WRITE: / 'Rows loaded from file:', lines( gt_upload ).
ENDFORM.

*----------------------------------------------------------------------*
* Form: Validate Data
*----------------------------------------------------------------------*
FORM f_validate_data USING iv_allow_ovpay TYPE xfeld.
  DATA: lv_row        TYPE i VALUE 0,
        lv_kna1_cnt   TYPE i,
        ls_bsid       TYPE bsid,
        lv_bkpf_cnt   TYPE i,
        lt_inv_split  TYPE TABLE OF string,
        lv_inv_token  TYPE string,
        lv_inv_ref    TYPE belnr_d,
        lv_total_open TYPE wrbtr,
        lv_inv_ok     TYPE i,
        lv_inv_err    TYPE xfeld.

  LOOP AT gt_upload INTO gs_upload.
    lv_row = lv_row + 1.
    CLEAR gs_log.
    gs_log-row_num        = lv_row.
    gs_log-company_code   = gs_upload-company_code.
    gs_log-customer_id    = gs_upload-customer_id.
    gs_log-invoice_refs   = gs_upload-invoice_refs.
    gs_log-payment_amount = gs_upload-payment_amount.
    gs_log-status         = 'PENDING'.

    " --- Validation 1: Required fields ---
    IF gs_upload-company_code IS INITIAL OR gs_upload-customer_id IS INITIAL
    OR gs_upload-invoice_refs IS INITIAL OR gs_upload-payment_amount IS INITIAL.
      gs_log-status  = 'ERROR'.
      gs_log-message = 'Missing required field(s): Company Code / Customer / Invoice(s) / Amount'.
      APPEND gs_log TO gt_log.
      CONTINUE.
    ENDIF.

    " --- Validation 2: Customer ID exists in KNA1 ---
    SELECT COUNT(*) FROM kna1 INTO lv_kna1_cnt
      WHERE kunnr = gs_upload-customer_id.
    IF lv_kna1_cnt = 0.
      gs_log-status  = 'ERROR'.
      gs_log-message = 'Customer ID Not Found in KNA1'.
      APPEND gs_log TO gt_log.
      CONTINUE.
    ENDIF.

    " --- Validation 3: Duplicate transaction ID ---
    IF gs_upload-transaction_id IS NOT INITIAL.
      SELECT COUNT(*) FROM bkpf INTO lv_bkpf_cnt
        WHERE bukrs = gs_upload-company_code
          AND xblnr = gs_upload-transaction_id.
      IF lv_bkpf_cnt > 0.
        gs_log-status  = 'ERROR'.
        gs_log-message = 'Duplicate Entry – Transaction ID already posted'.
        APPEND gs_log TO gt_log.
        CONTINUE.
      ENDIF.
    ENDIF.

    " --- Validation 4: Validate each invoice number individually ---
    "     Split semicolon-delimited invoice list and check BSID / BSAD.
    "     Each failed invoice gets its own log row; all failures are
    "     collected before the row is skipped so that the user sees every
    "     problem in a single run.
    CLEAR: lt_inv_split, lv_total_open, lv_inv_ok, lv_inv_err.
    SPLIT gs_upload-invoice_refs AT ';' INTO TABLE lt_inv_split.

    LOOP AT lt_inv_split INTO lv_inv_token.
      " Trim whitespace
      CONDENSE lv_inv_token NO-GAPS.
      IF lv_inv_token IS INITIAL.
        CONTINUE.
      ENDIF.

      lv_inv_ref = lv_inv_token.

      " Check open items (BSID)
      SELECT SINGLE * FROM bsid INTO ls_bsid
        WHERE bukrs = gs_upload-company_code
          AND kunnr = gs_upload-customer_id
          AND belnr = lv_inv_ref.
      IF sy-subrc <> 0.
        " Check already-cleared (BSAD)
        SELECT COUNT(*) FROM bsad INTO lv_kna1_cnt
          WHERE bukrs = gs_upload-company_code
            AND kunnr = gs_upload-customer_id
            AND belnr = lv_inv_ref.

        " Write one dedicated log row for this failed invoice
        DATA ls_inv_err_log TYPE ty_log.
        ls_inv_err_log = gs_log.
        ls_inv_err_log-invoice_refs = lv_inv_ref.
        ls_inv_err_log-status       = 'ERROR'.
        IF lv_kna1_cnt > 0.
          ls_inv_err_log-message = |Invoice { lv_inv_ref } Already Cleared (exists in BSAD)|.
        ELSE.
          ls_inv_err_log-message = |Invoice { lv_inv_ref } Not Found in open items (BSID)|.
        ENDIF.
        APPEND ls_inv_err_log TO gt_log.

        lv_inv_err = abap_true.
        CONTINUE.   " keep checking remaining invoices
      ENDIF.

      " Accumulate total open amount across all invoices
      ADD ls_bsid-wrbtr TO lv_total_open.
      ADD 1 TO lv_inv_ok.

      " Store validated invoice for posting step
      CLEAR gs_invoice.
      gs_invoice-row_num     = lv_row.
      gs_invoice-invoice_ref = lv_inv_ref.
      gs_invoice-open_amount = ls_bsid-wrbtr.
      gs_invoice-fiscal_year = ls_bsid-gjahr.
      APPEND gs_invoice TO gt_invoices.
    ENDLOOP.

    IF lv_inv_err = abap_true.
      " Per-invoice error rows already appended above; skip this upload row
      CONTINUE.
    ENDIF.

    IF lv_inv_ok = 0.
      gs_log-status  = 'ERROR'.
      gs_log-message = 'No valid invoice numbers found in Invoice(s) column'.
      APPEND gs_log TO gt_log.
      CONTINUE.
    ENDIF.

    " Store summed invoice amount back to upload row
    gs_upload-invoice_amount = lv_total_open.
    MODIFY gt_upload FROM gs_upload.

    " --- Validation 5: Payment vs total invoice amount ---
    IF gs_upload-payment_amount > lv_total_open.
      gs_log-residual_amt = gs_upload-payment_amount - lv_total_open.
      IF iv_allow_ovpay = 'X'.
        " Warning – excess will be credited on-account to customer
        gs_log-status  = 'PENDING'.
        gs_log-message = |{ lv_inv_ok } invoice(s) – overpayment warning, excess { gs_log-residual_amt } will be posted on-account to customer|.
      ELSE.
        " Error – overpayment rejected
        gs_log-status  = 'ERROR'.
        gs_log-message = |Payment { gs_upload-payment_amount } exceeds total open balance { lv_total_open } – rejected|.
        APPEND gs_log TO gt_log.
        CONTINUE.
      ENDIF.
    ELSEIF gs_upload-payment_amount < lv_total_open.
      gs_log-residual_amt = lv_total_open - gs_upload-payment_amount.
      gs_log-status  = 'PENDING'.
      gs_log-message = |{ lv_inv_ok } invoice(s) – partial payment, residual { gs_log-residual_amt } will be created|.
    ELSE.
      gs_log-status  = 'PENDING'.
      gs_log-message = |{ lv_inv_ok } invoice(s) validated – full payment match|.
    ENDIF.

    APPEND gs_log TO gt_log.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Post Payments
*----------------------------------------------------------------------*
FORM f_post_payments USING iv_test TYPE xfeld iv_allow_ovpay TYPE xfeld.
  DATA:
    ls_log            TYPE ty_log,
    lv_idx            TYPE sy-tabix,
    lv_row            TYPE i VALUE 0,
    lv_item_no        TYPE numc10,
    lv_item_ctr       TYPE i,

    " BAPI_ACC_DOCUMENT_POST structures
    ls_doc_header     TYPE acc_document_header,
    lt_account_gl     TYPE STANDARD TABLE OF accgl,
    ls_account_gl     TYPE accgl,
    lt_account_recv   TYPE STANDARD TABLE OF accreceivable,
    ls_account_recv   TYPE accreceivable,
    lt_open_items     TYPE STANDARD TABLE OF acc_open_item,
    ls_open_item      TYPE acc_open_item,
    lt_return         TYPE STANDARD TABLE OF bapiret2,
    ls_return         TYPE bapiret2,
    lv_doc_num        TYPE belnr_d,
    lv_obj_key        TYPE bapiplkkey,

    " Per-invoice residual distribution
    lt_row_invoices   TYPE STANDARD TABLE OF ty_invoice,
    ls_row_inv        TYPE ty_invoice,
    lv_remaining_pay  TYPE wrbtr,
    lv_inv_pay        TYPE wrbtr,
    lv_inv_residual   TYPE wrbtr.

  LOOP AT gt_upload INTO gs_upload.
    lv_row = lv_row + 1.

    " Find the corresponding log entry
    READ TABLE gt_log INTO ls_log
      WITH KEY row_num = lv_row.
    lv_idx = sy-tabix.

    " Skip rows that already failed validation
    IF ls_log-status = 'ERROR'.
      ADD 1 TO gv_total_err.
      CONTINUE.
    ENDIF.

    " Collect validated invoices for this row
    CLEAR lt_row_invoices.
    LOOP AT gt_invoices INTO ls_row_inv
      WHERE row_num = lv_row.
      APPEND ls_row_inv TO lt_row_invoices.
    ENDLOOP.

    " ---------------------------------------------------------------
    " Build BAPI_ACC_DOCUMENT_POST parameter structures
    " ---------------------------------------------------------------

    " Document Header
    CLEAR ls_doc_header.
    ls_doc_header-bus_act       = 'RFBU'.
    ls_doc_header-username      = sy-uname.
    ls_doc_header-comp_code     = gs_upload-company_code.
    ls_doc_header-doc_date      = gs_upload-payment_date.
    ls_doc_header-pstng_date    = gs_upload-posting_date.
    ls_doc_header-doc_type      = p_dtype.
    ls_doc_header-ref_doc_no    = gs_upload-transaction_id.
    ls_doc_header-header_txt    = gs_upload-text.
    ls_doc_header-currency      = gs_upload-currency.

    " G/L Line item 1 – Debit: Payment Processor Clearing Account (single line)
    CLEAR ls_account_gl.
    ls_account_gl-itemno_acc    = '0000000001'.
    ls_account_gl-gl_account    = gs_upload-gl_account.
    ls_account_gl-comp_code     = gs_upload-company_code.
    ls_account_gl-pstng_date    = gs_upload-posting_date.
    ls_account_gl-doc_type      = p_dtype.
    ls_account_gl-fisc_year     = gs_upload-posting_date(4).
    ls_account_gl-currency      = gs_upload-currency.
    ls_account_gl-amt_doccur    = gs_upload-payment_amount.   " Total debit (+)
    ls_account_gl-value_date    = gs_upload-value_date.
    ls_account_gl-item_text     = gs_upload-text.
    ls_account_gl-bank_acct     = gs_upload-house_bank_id.
    APPEND ls_account_gl TO lt_account_gl.

    " ---------------------------------------------------------------
    " One A/R credit line + one open-item clearing entry per invoice
    " Payment is distributed across invoices in order; last invoice
    " absorbs any residual if partial payment.
    " ---------------------------------------------------------------
    lv_item_ctr      = 1.
    lv_remaining_pay = gs_upload-payment_amount.

    LOOP AT lt_row_invoices INTO ls_row_inv.
      ADD 1 TO lv_item_ctr.
      lv_item_no = lv_item_ctr.

      " Determine how much of the payment applies to this invoice
      IF lv_remaining_pay >= ls_row_inv-open_amount.
        lv_inv_pay     = ls_row_inv-open_amount.  " Full invoice cleared
        lv_inv_residual = 0.
      ELSE.
        lv_inv_pay      = lv_remaining_pay.        " Partial – last invoice
        lv_inv_residual = ls_row_inv-open_amount - lv_remaining_pay.
      ENDIF.
      SUBTRACT lv_inv_pay FROM lv_remaining_pay.

      " A/R Customer credit line for this invoice
      CLEAR ls_account_recv.
      ls_account_recv-itemno_acc  = lv_item_no.
      ls_account_recv-customer    = gs_upload-customer_id.
      ls_account_recv-comp_code   = gs_upload-company_code.
      ls_account_recv-pstng_date  = gs_upload-posting_date.
      ls_account_recv-currency    = gs_upload-currency.
      ls_account_recv-amt_doccur  = lv_inv_pay * -1.   " Credit (-)
      ls_account_recv-bline_date  = gs_upload-payment_date.
      ls_account_recv-item_text   = gs_upload-text.
      APPEND ls_account_recv TO lt_account_recv.

      " Open item clearing entry – links credit line to specific invoice
      CLEAR ls_open_item.
      ls_open_item-itemno_acc   = lv_item_no.
      ls_open_item-op_item_type = 'D'.                " Debitor
      ls_open_item-comp_code    = gs_upload-company_code.
      ls_open_item-doc_no       = ls_row_inv-invoice_ref.
      ls_open_item-fisc_year    = ls_row_inv-fiscal_year.
      ls_open_item-currency     = gs_upload-currency.
      IF lv_inv_residual > 0.
        ls_open_item-pmnt_diff  = lv_inv_residual * -1.  " Residual (DF05B-PSDIF)
      ENDIF.
      APPEND ls_open_item TO lt_open_items.
    ENDLOOP.

    " ---------------------------------------------------------------
    " If payment exceeded total invoices and overpayment is allowed,
    " post excess as on-account credit on the customer
    " (no open-item clearing link).
    " ---------------------------------------------------------------
    IF lv_remaining_pay > 0 AND iv_allow_ovpay = 'X'.
      ADD 1 TO lv_item_ctr.
      lv_item_no = lv_item_ctr.

      CLEAR ls_account_recv.
      ls_account_recv-itemno_acc  = lv_item_no.
      ls_account_recv-customer    = gs_upload-customer_id.
      ls_account_recv-comp_code   = gs_upload-company_code.
      ls_account_recv-pstng_date  = gs_upload-posting_date.
      ls_account_recv-currency    = gs_upload-currency.
      ls_account_recv-amt_doccur  = lv_remaining_pay * -1.   " On-account credit (-)
      ls_account_recv-bline_date  = gs_upload-payment_date.
      ls_account_recv-item_text   = |On-account overpayment – { gs_upload-text }|.
      APPEND ls_account_recv TO lt_account_recv.
    ENDIF.

    " ---------------------------------------------------------------
    " Call BAPI_ACC_DOCUMENT_POST (Test or Live)
    " ---------------------------------------------------------------
    CLEAR: lt_return, lv_doc_num.

    IF iv_test = 'X'.
      CALL FUNCTION 'BAPI_ACC_DOCUMENT_POST'
        EXPORTING
          documentheader     = ls_doc_header
        IMPORTING
          obj_key            = lv_obj_key
        TABLES
          accountgl          = lt_account_gl
          accountreceivable  = lt_account_recv
          accountpayable     = lt_open_items
          return             = lt_return.

      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.   " Undo in test mode
      lv_doc_num = '[TEST]'.
    ELSE.
      CALL FUNCTION 'BAPI_ACC_DOCUMENT_POST'
        EXPORTING
          documentheader     = ls_doc_header
        IMPORTING
          obj_key            = lv_obj_key
        TABLES
          accountgl          = lt_account_gl
          accountreceivable  = lt_account_recv
          accountpayable     = lt_open_items
          return             = lt_return.

      READ TABLE lt_return INTO ls_return WITH KEY type = 'E'.
      IF sy-subrc = 0.
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      ELSE.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING wait = 'X'.
        lv_doc_num = lv_obj_key-obj_key(10).
      ENDIF.
    ENDIF.

    " ---------------------------------------------------------------
    " Update log entry with result
    " ---------------------------------------------------------------
    READ TABLE lt_return INTO ls_return WITH KEY type = 'E'.
    IF sy-subrc = 0.
      ls_log-status    = 'ERROR'.
      ls_log-message   = ls_return-message.
      ADD 1 TO gv_total_err.
    ELSE.
      IF iv_test = 'X'.
        ls_log-status  = 'SIM-OK'.
        IF lv_remaining_pay > 0 AND iv_allow_ovpay = 'X'.
          ls_log-message = |Simulation OK – { lines( lt_row_invoices ) } invoice(s) would be cleared, excess { lv_remaining_pay } posted on-account|.
        ELSE.
          ls_log-message = |Simulation OK – { lines( lt_row_invoices ) } invoice(s) would be cleared|.
        ENDIF.
      ELSE.
        ls_log-status  = 'SUCCESS'.
        IF lv_remaining_pay > 0 AND iv_allow_ovpay = 'X'.
          ls_log-message = |Document { lv_doc_num } posted – { lines( lt_row_invoices ) } invoice(s) cleared, excess { lv_remaining_pay } posted on-account|.
        ELSE.
          ls_log-message = |Document { lv_doc_num } posted – { lines( lt_row_invoices ) } invoice(s) cleared|.
        ENDIF.
        ADD 1 TO gv_total_ok.
        ADD gs_upload-payment_amount TO gv_total_amt.
      ENDIF.
      ls_log-sap_doc_num = lv_doc_num.
    ENDIF.

    MODIFY gt_log FROM ls_log INDEX lv_idx.

    " Clear work tables for next row
    CLEAR: ls_doc_header, lt_account_gl, lt_account_recv,
           lt_open_items, lt_return, ls_return.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Write Job Log / Spool Output
*   Writes the execution summary and every row result as plain WRITE
*   statements so the output lands in the spool/job log without opening
*   an ALV pop-up.  Suitable for background job execution.
*----------------------------------------------------------------------*
FORM f_write_job_log.
  DATA: lv_mode TYPE string.

  IF p_tst = 'X'.
    lv_mode = 'TEST MODE (simulation – no documents posted)'.
  ELSE.
    lv_mode = 'LIVE MODE'.
  ENDIF.

  WRITE: / '=== Mass Cash Posting – ', lv_mode, ' ==='.
  WRITE: / 'Run date/time :', sy-datum, sy-uzeit.
  WRITE: / 'Run by        :', sy-uname.
  WRITE: / 'Company code  :', p_bukrs.
  SKIP.
  WRITE: / '=== Execution Summary ==='.
  WRITE: / 'Successfully posted :', gv_total_ok.
  WRITE: / 'Errors / skipped    :', gv_total_err.
  WRITE: / 'Total amount posted :', gv_total_amt.
  SKIP.
  WRITE: / '=== Row-Level Results ==='.
  WRITE: /  3 'Row',
            8 'Customer',
           20 'Status',
           32 'SAP Doc #',
           44 'Message'.
  ULINE.

  LOOP AT gt_log INTO gs_log.
    WRITE: /  3 gs_log-row_num,
              8 gs_log-customer_id,
             20 gs_log-status,
             32 gs_log-sap_doc_num,
             44 gs_log-message.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Display ALV Results Log
*   1. ALV grid showing all parsed upload rows (file preview)
*   2. ALV grid showing posting log with traffic-light, icon, status,
*      and message columns
*----------------------------------------------------------------------*
FORM f_display_alv.

  " ---------------------------------------------------------------
  " Assign traffic_light and icon to every log entry before display
  " ---------------------------------------------------------------
  LOOP AT gt_log INTO gs_log.
    CASE gs_log-status.
      WHEN 'SUCCESS' OR 'SIM-OK'.
        gs_log-traffic_light = '3'.          " Green
        gs_log-icon          = icon_led_green.
      WHEN 'WARNING' OR 'PENDING'.
        gs_log-traffic_light = '2'.          " Yellow
        gs_log-icon          = icon_led_yellow.
      WHEN OTHERS.                            " ERROR
        gs_log-traffic_light = '1'.          " Red
        gs_log-icon          = icon_led_red.
    ENDCASE.
    MODIFY gt_log FROM gs_log.
  ENDLOOP.

  " Print summary to list
  WRITE: / '=== Execution Summary ==='.
  WRITE: / 'Successfully Posted:', gv_total_ok.
  WRITE: / 'Failed / Errors:    ', gv_total_err.
  WRITE: / 'Total Amount Posted:', gv_total_amt.
  SKIP.

  " ---------------------------------------------------------------
  " ALV 1: Uploaded file rows (preview of what was parsed)
  " ---------------------------------------------------------------
  DATA: lo_alv_up    TYPE REF TO cl_salv_table,
        lo_cols_up   TYPE REF TO cl_salv_columns_table,
        lo_col_up    TYPE REF TO cl_salv_column_table,
        lo_disp_up   TYPE REF TO cl_salv_display_settings,
        lo_funcs_up  TYPE REF TO cl_salv_functions_list.

  TRY.
    cl_salv_table=>factory(
      IMPORTING
        r_salv_table = lo_alv_up
      CHANGING
        t_table      = gt_upload ).
  CATCH cx_salv_msg.
    MESSAGE w001(00) WITH 'Error creating upload preview ALV'.
  ENDTRY.

  IF lo_alv_up IS BOUND.
    lo_funcs_up = lo_alv_up->get_functions( ).
    lo_funcs_up->set_all( abap_true ).

    lo_cols_up = lo_alv_up->get_columns( ).
    lo_cols_up->set_optimize( abap_true ).

    TRY.
      lo_col_up ?= lo_cols_up->get_column( 'COMPANY_CODE' ).
      lo_col_up->set_long_text( 'Company Code' ).

      lo_col_up ?= lo_cols_up->get_column( 'CUSTOMER_ID' ).
      lo_col_up->set_long_text( 'Customer ID' ).

      lo_col_up ?= lo_cols_up->get_column( 'INVOICE_REFS' ).
      lo_col_up->set_long_text( 'Invoice Number(s)' ).

      lo_col_up ?= lo_cols_up->get_column( 'PAYMENT_AMOUNT' ).
      lo_col_up->set_long_text( 'Payment Amount' ).

      lo_col_up ?= lo_cols_up->get_column( 'INVOICE_AMOUNT' ).
      lo_col_up->set_long_text( 'Invoice Amount' ).

      lo_col_up ?= lo_cols_up->get_column( 'PAYMENT_DATE' ).
      lo_col_up->set_long_text( 'Payment Date' ).

      lo_col_up ?= lo_cols_up->get_column( 'POSTING_DATE' ).
      lo_col_up->set_long_text( 'Posting Date' ).

      lo_col_up ?= lo_cols_up->get_column( 'VALUE_DATE' ).
      lo_col_up->set_long_text( 'Value Date' ).

      lo_col_up ?= lo_cols_up->get_column( 'CURRENCY' ).
      lo_col_up->set_long_text( 'Currency' ).

      lo_col_up ?= lo_cols_up->get_column( 'HOUSE_BANK' ).
      lo_col_up->set_long_text( 'House Bank' ).

      lo_col_up ?= lo_cols_up->get_column( 'HOUSE_BANK_ID' ).
      lo_col_up->set_long_text( 'House Bank Acct' ).

      lo_col_up ?= lo_cols_up->get_column( 'GL_ACCOUNT' ).
      lo_col_up->set_long_text( 'G/L Account' ).

      lo_col_up ?= lo_cols_up->get_column( 'TEXT' ).
      lo_col_up->set_long_text( 'Text' ).

      lo_col_up ?= lo_cols_up->get_column( 'TRANSACTION_ID' ).
      lo_col_up->set_long_text( 'Transaction ID' ).
    CATCH cx_salv_not_found.
      " Non-critical – continue
    ENDTRY.

    lo_disp_up = lo_alv_up->get_display_settings( ).
    lo_disp_up->set_striped_pattern( abap_true ).
    lo_disp_up->set_list_header( 'Uploaded File – Parsed Rows' ).

    lo_alv_up->display( ).
  ENDIF.

  " ---------------------------------------------------------------
  " ALV 2: Posting log with traffic light, icon, status, message
  " ---------------------------------------------------------------
  TRY.
    cl_salv_table=>factory(
      IMPORTING
        r_salv_table = go_alv
      CHANGING
        t_table      = gt_log ).
  CATCH cx_salv_msg.
    MESSAGE e001(00) WITH 'Error creating results log ALV'.
    RETURN.
  ENDTRY.

  " Enable standard ALV functions (sort, filter, export)
  go_funcs = go_alv->get_functions( ).
  go_funcs->set_all( abap_true ).

  " Set column headers and configure traffic-light / icon columns
  go_columns = go_alv->get_columns( ).
  go_columns->set_optimize( abap_true ).

  " Activate row-level traffic light using TRAFFIC_LIGHT field
  go_columns->set_exception_column( 'TRAFFIC_LIGHT' ).

  TRY.
    " Hide the raw traffic_light value column – it drives row colour only
    go_column ?= go_columns->get_column( 'TRAFFIC_LIGHT' ).
    go_column->set_visible( abap_false ).

    go_column ?= go_columns->get_column( 'ICON' ).
    go_column->set_long_text( 'Status' ).
    go_column->set_icon( abap_true ).

    go_column ?= go_columns->get_column( 'ROW_NUM' ).
    go_column->set_long_text( 'Row' ).

    go_column ?= go_columns->get_column( 'COMPANY_CODE' ).
    go_column->set_long_text( 'Company Code' ).

    go_column ?= go_columns->get_column( 'CUSTOMER_ID' ).
    go_column->set_long_text( 'Customer ID' ).

    go_column ?= go_columns->get_column( 'INVOICE_REFS' ).
    go_column->set_long_text( 'Invoice Number(s)' ).

    go_column ?= go_columns->get_column( 'PAYMENT_AMOUNT' ).
    go_column->set_long_text( 'Payment Amount' ).

    go_column ?= go_columns->get_column( 'RESIDUAL_AMT' ).
    go_column->set_long_text( 'Residual Amount' ).

    go_column ?= go_columns->get_column( 'SAP_DOC_NUM' ).
    go_column->set_long_text( 'SAP Document #' ).

    go_column ?= go_columns->get_column( 'STATUS' ).
    go_column->set_long_text( 'Result' ).

    go_column ?= go_columns->get_column( 'MESSAGE' ).
    go_column->set_long_text( 'Message' ).
  CATCH cx_salv_not_found.
    " Non-critical – continue
  ENDTRY.

  " Display settings
  go_display = go_alv->get_display_settings( ).
  go_display->set_striped_pattern( abap_true ).
  IF p_tst = 'X'.
    go_display->set_list_header( 'Mass Cash Posting – TEST MODE (no documents created)' ).
  ELSE.
    go_display->set_list_header( 'Mass Cash Posting – LIVE MODE' ).
  ENDIF.

  go_alv->display( ).
ENDFORM.

*----------------------------------------------------------------------*
* Form: Send Email Report via BCS (Business Communication Services)
*   – Plain-text body:  execution summary + one line per log row
*   – CSV attachment:   full posting log for spreadsheet analysis
*   – Sender:           running user (CL_SAPUSER_BCS)
*   – Recipient:        p_email / p_ename (SMTP address and display name from selection screen)
*----------------------------------------------------------------------*
FORM f_send_email.
  DATA:
    lo_bcs        TYPE REF TO cl_bcs,
    lo_doc        TYPE REF TO cl_document_bcs,
    lo_sender     TYPE REF TO cl_sapuser_bcs,
    lo_recipient  TYPE REF TO if_recipient_bcs,
    lo_addr       TYPE REF TO cl_cam_address_bcs,
    lv_subject    TYPE so_obj_des,
    lt_body       TYPE bcsy_text,
    ls_body       TYPE soli,
    lt_csv_lines  TYPE bcsy_text,
    ls_csv        TYPE soli,
    lv_xstr       TYPE xstring,
    lv_sent       TYPE os_boolean,
    lv_mode       TYPE string,
    lv_line       TYPE string.

  " ---------------------------------------------------------------
  " Build plain-text body
  " ---------------------------------------------------------------
  IF p_tst = 'X'.
    lv_mode = 'TEST MODE (simulation – no documents posted)'.
  ELSE.
    lv_mode = 'LIVE MODE'.
  ENDIF.

  ls_body-line = |Mass Cash Posting – { lv_mode }|.        APPEND ls_body TO lt_body.
  ls_body-line = |Run date/time : { sy-datum } { sy-uzeit }|. APPEND ls_body TO lt_body.
  ls_body-line = |Run by        : { sy-uname }|.           APPEND ls_body TO lt_body.
  ls_body-line = |Company code  : { p_bukrs }|.            APPEND ls_body TO lt_body.
  CLEAR ls_body. APPEND ls_body TO lt_body.
  ls_body-line = '=== Execution Summary ==='.              APPEND ls_body TO lt_body.
  ls_body-line = |Successfully posted : { gv_total_ok }|.  APPEND ls_body TO lt_body.
  ls_body-line = |Errors / skipped    : { gv_total_err }|. APPEND ls_body TO lt_body.
  ls_body-line = |Total amount posted : { gv_total_amt }|. APPEND ls_body TO lt_body.
  CLEAR ls_body. APPEND ls_body TO lt_body.
  ls_body-line = '=== Row-Level Results ==='.              APPEND ls_body TO lt_body.
  ls_body-line = 'Row | Customer   | Invoice(s)         | Amount       | Status     | SAP Doc# | Message'.
  APPEND ls_body TO lt_body.
  ls_body-line = '----+-----------+--------------------+--------------+------------+----------+---------'.
  APPEND ls_body TO lt_body.

  LOOP AT gt_log INTO gs_log.
    lv_line = |{ gs_log-row_num WIDTH = 3 } | { gs_log-customer_id WIDTH = 10 } | { gs_log-invoice_refs WIDTH = 19 } | { gs_log-payment_amount WIDTH = 13 } | { gs_log-status WIDTH = 10 } | { gs_log-sap_doc_num WIDTH = 9 } | { gs_log-message }|.
    ls_body-line = lv_line.
    APPEND ls_body TO lt_body.
  ENDLOOP.

  " ---------------------------------------------------------------
  " Build CSV attachment (header row + one data row per log entry)
  " ---------------------------------------------------------------
  ls_csv-line = 'Row,Company Code,Customer ID,Invoice(s),Payment Amount,Residual Amount,SAP Doc #,Status,Message'.
  APPEND ls_csv TO lt_csv_lines.

  LOOP AT gt_log INTO gs_log.
    " Enclose message in quotes to handle embedded commas
    lv_line = |{ gs_log-row_num },{ gs_log-company_code },{ gs_log-customer_id },{ gs_log-invoice_refs },{ gs_log-payment_amount },{ gs_log-residual_amt },{ gs_log-sap_doc_num },{ gs_log-status },"{ gs_log-message }"|.
    ls_csv-line = lv_line.
    APPEND ls_csv TO lt_csv_lines.
  ENDLOOP.

  " ---------------------------------------------------------------
  " Create BCS document (plain text body)
  " ---------------------------------------------------------------
  lv_subject = |Mass Cash Posting Results – { sy-datum }|.

  TRY.
    lo_doc = cl_document_bcs=>create_document(
               i_type    = 'RAW'
               i_text    = lt_body
               i_subject = lv_subject ).

    " Convert CSV lines to xstring and attach
    CALL FUNCTION 'SCMS_TEXT_TO_XSTRING'
      EXPORTING
        mimetype  = 'text/csv'
        encoding  = 'UTF-8'
      IMPORTING
        xstring   = lv_xstr
      TABLES
        text_tab  = lt_csv_lines
      EXCEPTIONS
        failed    = 1
        OTHERS    = 2.

    IF sy-subrc = 0 AND lv_xstr IS NOT INITIAL.
      lo_doc->add_attachment(
        i_attachment_type    = 'CSV'
        i_attachment_subject = |MassCashPosting_{ sy-datum }.csv|
        i_att_content_hex    = lv_xstr ).
    ENDIF.

    " ---------------------------------------------------------------
    " Create send request and set sender / recipient
    " ---------------------------------------------------------------
    lo_bcs = cl_bcs=>create_persistent( ).

    lo_sender = cl_sapuser_bcs=>create( sy-uname ).
    lo_bcs->set_sender( lo_sender ).

    lo_addr = cl_cam_address_bcs=>create_internet_address(
                i_address_string = p_email
                i_address_name   = p_ename ).
    lo_recipient ?= lo_addr.
    lo_bcs->add_recipient(
      i_recipient = lo_recipient
      i_express   = abap_true ).

    lo_bcs->set_document( lo_doc ).

    lv_sent = lo_bcs->send( i_with_error_screen = abap_true ).

    IF lv_sent = abap_true.
      MESSAGE s001(00) WITH |Email report sent to { p_email }|.
    ELSE.
      MESSAGE w001(00) WITH |Email could not be sent to { p_email } – check SCOT config|.
    ENDIF.

  CATCH cx_bcs INTO DATA(lx_bcs).
    MESSAGE w001(00) WITH |BCS error sending email: { lx_bcs->get_text( ) }|.
  ENDTRY.
ENDFORM.
