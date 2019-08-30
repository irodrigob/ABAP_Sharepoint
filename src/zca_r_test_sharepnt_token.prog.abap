*&---------------------------------------------------------------------*
*& Report ZCA_R_TEST_SHAREPNT_TOKEN
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zca_r_test_sharepnt_token.

PARAMETERS: p_cl_id TYPE string LOWER CASE,
            p_cl_se TYPE string LOWER CASE,
            p_realm TYPE string LOWER CASE,
            p_host  TYPE string LOWER CASE.

START-OF-SELECTION.

  TRY.
      DATA(lv_token) = NEW zcl_ca_microsoft_auth( )->generate_token_auth( iv_client_secret = p_cl_se
                                                                          iv_client_id = p_cl_id
                                                                          iv_domain = p_host
                                                                          iv_realm = p_realm ).
      cl_demo_output=>display( lv_token ).

    CATCH zcx_ca_http_services INTO DATA(lo_excep).
      cl_demo_output=>write_text( text = |Status code: { lo_excep->mv_status_code }| ).
      cl_demo_output=>write_text( text = |Status message: { lo_excep->mv_status_text }| ).
      cl_demo_output=>write_text( text = |Response:| ).
      cl_demo_output=>write_json( json = lo_excep->mv_content_response ).
      cl_demo_output=>display( ).

  ENDTRY.
