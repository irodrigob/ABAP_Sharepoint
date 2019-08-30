CLASS zcl_ca_microsoft_auth DEFINITION
  PUBLIC
  INHERITING FROM zcl_ca_http_services
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS generate_token_auth
      IMPORTING iv_domain        TYPE any
                iv_client_id     TYPE any
                iv_client_secret TYPE any
                iv_realm         TYPE any
      RETURNING VALUE(rv_token)  TYPE string
      RAISING   zcx_ca_http_services.

  PROTECTED SECTION.
    TYPES: BEGIN OF ts_body_receive_access_token,
             token_type   TYPE string,
             expires_in   TYPE string,
             not_before   TYPE string,
             expires_on   TYPE string,
             resource     TYPE string,
             access_token TYPE string,
           END OF ts_body_receive_access_token.

    CONSTANTS mv_resource TYPE string VALUE '00000003-0000-0ff1-ce00-000000000000'.
    CONSTANTS cv_url_access_control TYPE string VALUE 'accounts.accesscontrol.windows.net'.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_ca_microsoft_auth IMPLEMENTATION.


  METHOD generate_token_auth.
    DATA ls_body_receive_access_token TYPE ts_body_receive_access_token.

    CLEAR: rv_token.

    create_http_client( iv_client = cv_url_access_control iv_is_https = abap_true ).

    set_form_value( iv_name  = 'grant_type' iv_value = 'client_credentials' ).
    set_form_value( iv_name  = 'client_id' iv_value = |{ iv_client_id }@{ iv_realm }| ).
    set_form_value( iv_name  = 'client_secret' iv_value = iv_client_secret ).
    set_form_value( iv_name  = 'resource' iv_value = |{ mv_resource }/{ iv_domain }@{ iv_realm }| ).

    " Tipo form urlencoded
    set_content_type( 'X' ).

    " El método POST
    set_request_method( 'POST' ).

    " Se indica a que URL se va a llamar
    set_request_uri( |/{ iv_realm }/tokens/OAuth/2| ).

    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-low_case
             IMPORTING ev_data        = ls_body_receive_access_token ).

    " El token es la concatenación de dos campos y se guarda en una variable global y se devolverá por parámetro
    rv_token = |{ ls_body_receive_access_token-token_type } { ls_body_receive_access_token-access_token }|.

  ENDMETHOD.
ENDCLASS.
