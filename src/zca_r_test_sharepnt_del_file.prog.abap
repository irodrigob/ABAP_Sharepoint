*&---------------------------------------------------------------------*
*& Report ZCA_R_TEST_SHAREPNT_DEL_FILE
*&--------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zca_r_test_sharepnt_del_file.


PARAMETERS: p_cl_id TYPE string LOWER CASE,
            p_cl_se TYPE string LOWER CASE,
            p_realm TYPE string LOWER CASE,
            p_host  TYPE string LOWER CASE,
            p_site  TYPE string LOWER CASE,
            p_dlib  TYPE string LOWER CASE,
            p_https AS CHECKBOX DEFAULT 'X'.

SELECTION-SCREEN SKIP 2.

PARAMETERS p_folder TYPE string LOWER CASE.
PARAMETERS p_file TYPE string LOWER CASE.
PARAMETERS p_check AS CHECKBOX.

START-OF-SELECTION.



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

      mo_sharepoint->delete_file(
        EXPORTING
          iv_path_folder  =  p_folder
          iv_file_name    = p_file
          iv_verify_delete = p_check
        IMPORTING
          ev_file_deleted  = DATA(lv_file_deleted) ).

      IF lv_file_deleted = abap_true.
        WRITE:/ 'File deleted'.
      ELSE.
        WRITE:/ 'Errot to delete the file'.
      ENDIF.

    CATCH zcx_ca_http_services INTO DATA(lo_excep).
      cl_demo_output=>write_text( text = |Status code: { lo_excep->mv_status_code }| ).
      cl_demo_output=>write_text( text = |Status message: { lo_excep->mv_status_text }| ).
      cl_demo_output=>write_text( text = |Response:| ).
      cl_demo_output=>write_html( html = lo_excep->mv_content_response ).
      cl_demo_output=>display( ).
  ENDTRY.
