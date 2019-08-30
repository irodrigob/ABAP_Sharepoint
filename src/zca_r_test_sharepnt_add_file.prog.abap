*&---------------------------------------------------------------------*
*& Report ZCA_R_TEST_SHAREPNT_ADD_FILE
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zca_r_test_sharepnt_add_file.

DATA mt_filetable TYPE filetable.
DATA mt_file_bin TYPE solix_tab.
DATA mv_file_content TYPE xstring.
DATA mv_file_name TYPE string.


PARAMETERS: p_cl_id TYPE string LOWER CASE,
            p_cl_se TYPE string LOWER CASE,
            p_realm TYPE string LOWER CASE,
            p_host  TYPE string LOWER CASE,
            p_site  TYPE string LOWER CASE,
            p_dlib  TYPE string LOWER CASE,
            p_https AS CHECKBOX DEFAULT 'X'.

SELECTION-SCREEN SKIP 2.

PARAMETERS p_folder TYPE string LOWER CASE.
PARAMETERS p_file TYPE rlgrap-filename.

    AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
      DATA lv_rc TYPE i.
      CALL METHOD cl_gui_frontend_services=>file_open_dialog
        EXPORTING
          multiselection = abap_false
          file_filter    = '*.*'
        CHANGING
          file_table     = mt_filetable
          rc             = lv_rc.

      READ TABLE mt_filetable INTO p_file INDEX 1.

START-OF-SELECTION.

* " Upload file
  cl_gui_frontend_services=>gui_upload(
     EXPORTING
       filename                = CONV string( p_file )
       filetype                = 'BIN'
  IMPORTING
       filelength              = DATA(mv_len)
     CHANGING
       data_tab                = mt_file_bin
     EXCEPTIONS
       file_open_error         = 1
       file_read_error         = 2
       no_batch                = 3
       gui_refuse_filetransfer = 4
       invalid_type            = 5
       no_authority            = 6
       unknown_error           = 7
       bad_data_format         = 8
       header_not_allowed      = 9
       separator_not_allowed   = 10
       header_too_long         = 11
       unknown_dp_error        = 12
       access_denied           = 13
       dp_out_of_memory        = 14
       disk_full               = 15
       dp_timeout              = 16
       not_supported_by_gui    = 17
       error_no_gui            = 18
       OTHERS                  = 19
   ).
  IF sy-subrc <> 0.
    WRITE:/ 'Error to read the file'.
  ELSE.


    " Convert binary file to xstring
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = mv_len
      IMPORTING
        buffer       = mv_file_content
      TABLES
        binary_tab   = mt_file_bin
      EXCEPTIONS
        failed       = 1
        OTHERS       = 2.

    " Get the name of file
    CALL FUNCTION 'SO_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = p_file
      IMPORTING
        stripped_name = mv_file_name.

    TRY.
        " Se instancia la clase de sharepoint
        DATA(mo_sharepoint) = NEW zcl_ca_sharepoint_rest(  iv_client_secret = p_cl_se
                                                           iv_client_id = p_cl_id
                                                           iv_domain = p_host
                                                           iv_realm = p_realm
                                                           iv_site = p_site
                                                           iv_document_library = p_dlib
                                                           iv_protocol_https = p_https
                                                           iv_generate_tokeh_auth = abap_true ).

        mo_sharepoint->add_file(
          EXPORTING
            iv_path_folder  =  p_folder
          iv_content_file = mv_file_content
          iv_file_name    = mv_file_name
        IMPORTING
          ev_file_upload  = DATA(lv_file_upload)
        ).

        IF lv_file_upload = abap_true.
          WRITE:/ 'File added'.
        ELSE.
          WRITE:/ 'Errot to upload the file'.
        ENDIF.

      CATCH zcx_ca_http_services INTO DATA(lo_excep).
        cl_demo_output=>write_text( text = |Status code: { lo_excep->mv_status_code }| ).
        cl_demo_output=>write_text( text = |Status message: { lo_excep->mv_status_text }| ).
        cl_demo_output=>write_text( text = |Response:| ).
        cl_demo_output=>write_json( json = lo_excep->mv_content_response ).
        cl_demo_output=>display( ).
    ENDTRY.

  ENDIF.
