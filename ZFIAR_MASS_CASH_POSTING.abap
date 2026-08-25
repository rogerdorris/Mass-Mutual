*&---------------------------------------------------------------------*
*& Report  ZFIAR_MASS_CASH_POSTING
*& Description: Mass Upload of Customer Payment and Clearing
*&              Emulates F-28 via POSTING_INTERFACE_CLEARING /
*&              POSTING_INTERFACE_DOCUMENT for bulk A/R cash posting.
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
    invoice_ref    TYPE belnr_d,     " RFOPS_DK-BELNR
    payment_amount TYPE wrbtr,       " BSEG-WRBTR
    invoice_amount TYPE wrbtr,       " Open item amount from BSID
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

  " Execution log line structure
  BEGIN OF ty_log,
    row_num        TYPE i,
    company_code   TYPE bukrs,
    customer_id    TYPE kunnr,
    invoice_ref    TYPE belnr_d,
    payment_amount TYPE wrbtr,
    residual_amt   TYPE wrbtr,
    sap_doc_num    TYPE belnr_d,
    status         TYPE char10,      " SUCCESS / WARNING / ERROR
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
  gs_raw      TYPE alsmex_tabline.

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
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
PARAMETERS:
  p_file   TYPE string OBLIGATORY,                " Local file path
  p_bukrs  TYPE bukrs DEFAULT '2920',             " Company code filter
  p_test   TYPE xfeld DEFAULT 'X'.                " Test mode flag
SELECTION-SCREEN END OF BLOCK b1.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f_browse_file CHANGING p_file.

*----------------------------------------------------------------------*
* Main Program
*----------------------------------------------------------------------*
START-OF-SELECTION.

  " Authority check
  PERFORM f_authority_check.

  " Upload and parse the Excel file
  PERFORM f_upload_file USING p_file.

  " Validate all rows before posting
  PERFORM f_validate_data.

  " Post or simulate
  PERFORM f_post_payments USING p_test.

  " Display ALV log
  PERFORM f_display_alv.

*----------------------------------------------------------------------*
* Form: Browse File (F4 help for file path)
*----------------------------------------------------------------------*
FORM f_browse_file CHANGING cv_file TYPE string.
  DATA: lv_rc      TYPE i,
        lt_filetab TYPE filetable,
        ls_filetab LIKE LINE OF lt_filetab.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Select Cash Upload File'
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
*----------------------------------------------------------------------*
FORM f_upload_file USING iv_file TYPE string.
  DATA: lv_filename TYPE string.

  lv_filename = iv_file.

  " Read Excel into raw internal table (rows/columns)
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
    MESSAGE e001(00) WITH 'Error reading Excel file. Check path and format.'.
    STOP.
  ENDIF.

  IF gt_raw IS INITIAL.
    MESSAGE e001(00) WITH 'No data found in upload file.'.
    STOP.
  ENDIF.

  " Map raw cells to typed upload structure
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
      WHEN 3.  gs_upload-invoice_ref    = gs_raw-value.
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

  WRITE: / 'Rows loaded from file:', lines( gt_upload ).
ENDFORM.

*----------------------------------------------------------------------*
* Form: Validate Data
*----------------------------------------------------------------------*
FORM f_validate_data.
  DATA: lv_row       TYPE i VALUE 0,
        lv_kna1_cnt  TYPE i,
        ls_bsid      TYPE bsid,
        lv_bkpf_cnt  TYPE i.

  LOOP AT gt_upload INTO gs_upload.
    lv_row = lv_row + 1.
    CLEAR gs_log.
    gs_log-row_num        = lv_row.
    gs_log-company_code   = gs_upload-company_code.
    gs_log-customer_id    = gs_upload-customer_id.
    gs_log-invoice_ref    = gs_upload-invoice_ref.
    gs_log-payment_amount = gs_upload-payment_amount.
    gs_log-status         = 'PENDING'.

    " --- Validation 1: Required fields ---
    IF gs_upload-company_code IS INITIAL OR gs_upload-customer_id IS INITIAL
    OR gs_upload-invoice_ref  IS INITIAL OR gs_upload-payment_amount IS INITIAL.
      gs_log-status  = 'ERROR'.
      gs_log-message = 'Missing required field(s): Company Code / Customer / Invoice / Amount'.
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

    " --- Validation 3: Open invoice exists in BSID ---
    SELECT SINGLE * FROM bsid INTO ls_bsid
      WHERE bukrs = gs_upload-company_code
        AND kunnr = gs_upload-customer_id
        AND belnr = gs_upload-invoice_ref.
    IF sy-subrc <> 0.
      " Check if already cleared (exists in BSAD)
      SELECT COUNT(*) FROM bsad INTO lv_kna1_cnt
        WHERE bukrs = gs_upload-company_code
          AND kunnr = gs_upload-customer_id
          AND belnr = gs_upload-invoice_ref.
      IF lv_kna1_cnt > 0.
        gs_log-status  = 'ERROR'.
        gs_log-message = 'Invoice Already Cleared (exists in BSAD)'.
      ELSE.
        gs_log-status  = 'ERROR'.
        gs_log-message = 'Invoice Not Found in open items (BSID)'.
      ENDIF.
      APPEND gs_log TO gt_log.
      CONTINUE.
    ENDIF.

    " Store invoice open amount for residual calculation
    gs_upload-invoice_amount = ls_bsid-wrbtr.
    MODIFY gt_upload FROM gs_upload.

    " --- Validation 4: Duplicate transaction ID ---
    IF gs_upload-transaction_id IS NOT INITIAL.
      SELECT COUNT(*) FROM bkpf INTO lv_bkpf_cnt
        WHERE bukrs  = gs_upload-company_code
          AND xblnr  = gs_upload-transaction_id.
      IF lv_bkpf_cnt > 0.
        gs_log-status  = 'ERROR'.
        gs_log-message = 'Duplicate Entry – Transaction ID already posted'.
        APPEND gs_log TO gt_log.
        CONTINUE.
      ENDIF.
    ENDIF.

    " --- Validation 5: Payment vs invoice amount ---
    IF gs_upload-payment_amount > gs_upload-invoice_amount.
      gs_log-status  = 'ERROR'.
      gs_log-message = 'Payment Amount exceeds open invoice balance'.
      APPEND gs_log TO gt_log.
      CONTINUE.
    ELSEIF gs_upload-payment_amount < gs_upload-invoice_amount.
      gs_log-residual_amt = gs_upload-invoice_amount - gs_upload-payment_amount.
      gs_log-status  = 'PENDING'.
      gs_log-message = 'Partial payment – residual item will be created'.
    ELSE.
      gs_log-status  = 'PENDING'.
      gs_log-message = 'Validation OK – full payment match'.
    ENDIF.

    APPEND gs_log TO gt_log.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Post Payments
*----------------------------------------------------------------------*
FORM f_post_payments USING iv_test TYPE xfeld.
  DATA:
    ls_log            TYPE ty_log,
    lv_idx            TYPE sy-tabix,
    lv_row            TYPE i VALUE 0,

    " POSTING_INTERFACE_CLEARING structures
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
    lv_obj_key        TYPE bapiplkkey.

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

    " ---------------------------------------------------------------
    " Build BAPI_ACC_DOCUMENT_POST parameter structures
    " ---------------------------------------------------------------

    " Document Header
    CLEAR ls_doc_header.
    ls_doc_header-bus_act       = 'RFBU'.       " FI posting activity
    ls_doc_header-username      = sy-uname.
    ls_doc_header-comp_code     = gs_upload-company_code.
    ls_doc_header-doc_date      = gs_upload-payment_date.
    ls_doc_header-pstng_date    = gs_upload-posting_date.
    ls_doc_header-doc_type      = 'DZ'.
    ls_doc_header-ref_doc_no    = gs_upload-transaction_id.
    ls_doc_header-header_txt    = gs_upload-text.
    ls_doc_header-currency      = gs_upload-currency.

    " G/L Line (Debit: Payment Processor Clearing Account)
    CLEAR ls_account_gl.
    ls_account_gl-itemno_acc    = '0000000001'.
    ls_account_gl-gl_account    = gs_upload-gl_account.
    ls_account_gl-comp_code     = gs_upload-company_code.
    ls_account_gl-pstng_date    = gs_upload-posting_date.
    ls_account_gl-doc_type      = 'DZ'.
    ls_account_gl-fisc_year     = gs_upload-posting_date(4).
    ls_account_gl-currency      = gs_upload-currency.
    ls_account_gl-amt_doccur    = gs_upload-payment_amount.   " Debit (+)
    ls_account_gl-value_date    = gs_upload-value_date.
    ls_account_gl-item_text     = gs_upload-text.
    ls_account_gl-bank_acct     = gs_upload-house_bank_id.
    APPEND ls_account_gl TO lt_account_gl.

    " A/R Customer Line (Credit: clears open item)
    CLEAR ls_account_recv.
    ls_account_recv-itemno_acc  = '0000000002'.
    ls_account_recv-customer    = gs_upload-customer_id.
    ls_account_recv-comp_code   = gs_upload-company_code.
    ls_account_recv-pstng_date  = gs_upload-posting_date.
    ls_account_recv-currency    = gs_upload-currency.
    ls_account_recv-amt_doccur  = gs_upload-payment_amount * -1.  " Credit (-)
    ls_account_recv-pmnttrms    = ''.
    ls_account_recv-bline_date  = gs_upload-payment_date.
    ls_account_recv-item_text   = gs_upload-text.
    APPEND ls_account_recv TO lt_account_recv.

    " Open Item to Clear
    CLEAR ls_open_item.
    ls_open_item-itemno_acc     = '0000000002'.
    ls_open_item-op_item_type   = 'D'.           " Debitor
    ls_open_item-comp_code      = gs_upload-company_code.
    ls_open_item-doc_no         = gs_upload-invoice_ref.
    ls_open_item-fisc_year      = gs_upload-posting_date(4).
    ls_open_item-currency       = gs_upload-currency.

    " Residual handling: if partial payment, set residual amount
    IF ls_log-residual_amt > 0.
      ls_open_item-pmnt_diff    = ls_log-residual_amt * -1.   " Residual (DF05B-PSDIF equivalent)
    ENDIF.
    APPEND ls_open_item TO lt_open_items.

    " ---------------------------------------------------------------
    " Call BAPI_ACC_DOCUMENT_POST (Test or Live)
    " ---------------------------------------------------------------
    CLEAR: lt_return, lv_doc_num.

    IF iv_test = 'X'.
      " Simulation – call with SIMULATION = 'X' via BAPI_TRANSACTION_ROLLBACK after check
      CALL FUNCTION 'BAPI_ACC_DOCUMENT_POST'
        EXPORTING
          documentheader     = ls_doc_header
        IMPORTING
          obj_key            = lv_obj_key
        TABLES
          accountgl          = lt_account_gl
          accountreceivable  = lt_account_recv
          accountpayable     = lt_open_items   " reused for open item reference
          return             = lt_return.

      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.   " Undo in test mode

      lv_doc_num = '[TEST]'.
    ELSE.
      " Live posting
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

      " Check return for errors
      READ TABLE lt_return INTO ls_return
        WITH KEY type = 'E'.
      IF sy-subrc = 0.
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      ELSE.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = 'X'.
        lv_doc_num = lv_obj_key-obj_key(10).
      ENDIF.
    ENDIF.

    " ---------------------------------------------------------------
    " Update log entry with result
    " ---------------------------------------------------------------
    READ TABLE lt_return INTO ls_return
      WITH KEY type = 'E'.
    IF sy-subrc = 0.
      ls_log-status    = 'ERROR'.
      ls_log-message   = ls_return-message.
      ADD 1 TO gv_total_err.
    ELSE.
      IF iv_test = 'X'.
        ls_log-status  = 'SIM-OK'.
        ls_log-message = 'Simulation successful – no document created'.
      ELSE.
        ls_log-status  = 'SUCCESS'.
        ls_log-message = 'Document posted: ' && lv_doc_num.
        ADD 1 TO gv_total_ok.
        ADD gs_upload-payment_amount TO gv_total_amt.
      ENDIF.
      ls_log-sap_doc_num = lv_doc_num.
    ENDIF.

    MODIFY gt_log FROM ls_log INDEX lv_idx.

    " Clear work tables for next iteration
    CLEAR: ls_doc_header, lt_account_gl, lt_account_recv,
           lt_open_items, lt_return, ls_return.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form: Display ALV Results Log
*----------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lv_mode TYPE c.

  " Print summary to log
  WRITE: / '=== Execution Summary ==='.
  WRITE: / 'Successfully Posted:', gv_total_ok.
  WRITE: / 'Failed / Errors:    ', gv_total_err.
  WRITE: / 'Total Amount Posted:', gv_total_amt.
  SKIP.

  " Create ALV display
  TRY.
    cl_salv_table=>factory(
      IMPORTING
        r_salv_table = go_alv
      CHANGING
        t_table      = gt_log ).
  CATCH cx_salv_msg.
    MESSAGE e001(00) WITH 'Error creating ALV display'.
    RETURN.
  ENDTRY.

  " Enable standard ALV functions (sort, filter, export)
  go_funcs = go_alv->get_functions( ).
  go_funcs->set_all( abap_true ).

  " Set column headers
  go_columns = go_alv->get_columns( ).
  go_columns->set_optimize( abap_true ).

  TRY.
    go_column ?= go_columns->get_column( 'ROW_NUM' ).
    go_column->set_long_text( 'Row' ).

    go_column ?= go_columns->get_column( 'COMPANY_CODE' ).
    go_column->set_long_text( 'Company Code' ).

    go_column ?= go_columns->get_column( 'CUSTOMER_ID' ).
    go_column->set_long_text( 'Customer ID' ).

    go_column ?= go_columns->get_column( 'INVOICE_REF' ).
    go_column->set_long_text( 'Invoice Reference' ).

    go_column ?= go_columns->get_column( 'PAYMENT_AMOUNT' ).
    go_column->set_long_text( 'Payment Amount' ).

    go_column ?= go_columns->get_column( 'RESIDUAL_AMT' ).
    go_column->set_long_text( 'Residual Amount' ).

    go_column ?= go_columns->get_column( 'SAP_DOC_NUM' ).
    go_column->set_long_text( 'SAP Document #' ).

    go_column ?= go_columns->get_column( 'STATUS' ).
    go_column->set_long_text( 'Status' ).

    go_column ?= go_columns->get_column( 'MESSAGE' ).
    go_column->set_long_text( 'Message' ).
  CATCH cx_salv_not_found.
    " Non-critical – continue
  ENDTRY.

  " Display settings
  go_display = go_alv->get_display_settings( ).
  go_display->set_striped_pattern( abap_true ).
  IF p_test = 'X'.
    go_display->set_list_header( 'Mass Cash Posting – TEST MODE (no documents created)' ).
  ELSE.
    go_display->set_list_header( 'Mass Cash Posting – LIVE MODE' ).
  ENDIF.

  go_alv->display( ).
ENDFORM.
