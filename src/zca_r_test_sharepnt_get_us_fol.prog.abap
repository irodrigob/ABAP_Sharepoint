*&---------------------------------------------------------------------*
*& Report ZCA_R_TEST_SHAREPNT_GET_US_FOL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zca_r_test_sharepnt_get_us_fol.

PARAMETERS: p_cl_id TYPE string LOWER CASE,
            p_cl_se TYPE string LOWER CASE,
            p_realm TYPE string LOWER CASE,
            p_host  TYPE string LOWER CASE,
            p_site  TYPE string LOWER CASE,
            p_dlib  TYPE string LOWER CASE,
            p_https AS CHECKBOX DEFAULT 'X'.

SELECTION-SCREEN SKIP 2.

PARAMETERS: p_folder TYPE string LOWER CASE.


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

      mo_sharepoint->get_user_folder(
        EXPORTING
          iv_path_folder = p_folder
          iv_get_role = abap_true
        IMPORTING
          et_users = DATA(lt_users) ).

      cl_demo_output=>display( lt_users ).

    CATCH zcx_ca_http_services INTO DATA(lo_excep).
      cl_demo_output=>write_text( text = |Status code: { lo_excep->mv_status_code }| ).
      cl_demo_output=>write_text( text = |Status message: { lo_excep->mv_status_text }| ).
      cl_demo_output=>write_text( text = |Response:| ).
      cl_demo_output=>write_json( json = lo_excep->mv_content_response ).
      cl_demo_output=>display( ).
  ENDTRY.
