CLASS zcl_ca_sharepoint_rest DEFINITION
  PUBLIC
  INHERITING FROM zcl_ca_http_services
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    CONSTANTS: BEGIN OF cs_base_template,
                 document_library TYPE i VALUE 101,
               END OF cs_base_template.
    CONSTANTS: BEGIN OF cs_members,
                 BEGIN OF principal_type,
                   user  TYPE int4 VALUE 1,
                   group TYPE int4 VALUE 8,
                 END OF principal_type,
               END OF cs_members.
    TYPES:
      BEGIN OF ts_receive_get_folders_values,
        exists              TYPE string,
        is_wopi_enabled     TYPE string,
        item_count          TYPE i,
        name                TYPE string,
        prog_id             TYPE string,
        server_relative_url TYPE string,
        time_created        TYPE string,
        time_last_modified  TYPE string,
        unique_id           TYPE string,
        welcome_page        TYPE string,
      END OF ts_receive_get_folders_values .
    TYPES:
      tt_receive_get_folders_values TYPE STANDARD TABLE OF ts_receive_get_folders_values WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_receive_get_files_values,
        check_in_comment       TYPE string,
        check_out_type         TYPE i,
        content_tag            TYPE string,
        customized_page_status TYPE i,
        e_tag                  TYPE string,
        exists                 TYPE sap_bool,
        irm_enabled            TYPE sap_bool,
        length                 TYPE string,
        level                  TYPE i,
        linking_uri            TYPE string,
        linking_url            TYPE string,
        major_version          TYPE i,
        minor_version          TYPE i,
        name                   TYPE string,
        server_relative_url    TYPE string,
        time_created           TYPE string,
        time_last_modified     TYPE string,
        title                  TYPE string,
        ui_version             TYPE i,
        ui_version_label       TYPE string,
        uniqueid               TYPE string,
      END OF ts_receive_get_files_values .
    TYPES:
      tt_receive_get_files_values TYPE STANDARD TABLE OF ts_receive_get_files_values WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_receive_get_user_info,
        id                           TYPE int4,
        is_hidden_in_ui              TYPE sap_bool,
        login_name                   TYPE string,
        title                        TYPE string,
        principal_type               TYPE int4,
        email                        TYPE string,
        expiration                   TYPE string,
*             is_email_authentication_guest_user TYPE sap_bool,
        is_share_by_email_guest_user TYPE sap_bool,
        is_site_admin                TYPE sap_bool,
        BEGIN OF user_id,
          name_id        TYPE string,
          name_id_issuer TYPE string,
        END OF user_id,
        user_principal_name          TYPE string,
      END OF ts_receive_get_user_info .
    TYPES:
      BEGIN OF ts_get_user_folder_role,
        id   TYPE int4,
        name TYPE string,
      END OF ts_get_user_folder_role .
    TYPES:
      tt_get_user_folder_role TYPE STANDARD TABLE OF ts_get_user_folder_role WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_get_user_folder,
        user_id        TYPE int4,
        user_loginname TYPE string,
        user_type      TYPE int4,
        role           TYPE tt_get_user_folder_role,
      END OF ts_get_user_folder .
    TYPES:
      tt_get_user_folder TYPE STANDARD TABLE OF ts_get_user_folder WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_recieve_sitegroup_id,
        id          TYPE int4,
        login_name  TYPE string,
        title       TYPE string,
        description TYPE string,
        owner_title TYPE string,
      END OF ts_recieve_sitegroup_id .
    TYPES:
      tt_recieve_sitegroup_id TYPE STANDARD TABLE OF ts_recieve_sitegroup_id WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_get_user_folder_member,
        id                  TYPE int4,
        login_name          TYPE string,
        title               TYPE string,
        principal_type      TYPE int2,
        email               TYPE string,
        user_principal_name TYPE string,
        BEGIN OF user_id,
          name_id        TYPE string,
          name_id_issuer TYPE string,
        END OF user_id,
      END OF ts_get_user_folder_member .
    TYPES:
      BEGIN OF ts_receive_get_info_doc_lib,
        allow_content_types     TYPE string,
        base_template           TYPE i,
        base_type               TYPE i,
        content_types_enabled   TYPE string,
        crawl_non_default_views TYPE string,
        created                 TYPE string,
        timecreated             TYPE string,
        current_change_token    TYPE string,
        disablegridediting      TYPE string,
        document_template_url   TYPE string,
        id                      TYPE string,
        title                   TYPE string,
      END OF ts_receive_get_info_doc_lib .
    CONSTANTS:
      BEGIN OF cs_id_roles,
        full_control    TYPE int4 VALUE '1073741829',
        design          TYPE int4 VALUE '1073741828',
        edit            TYPE int4 VALUE '1073741830',
        contribute      TYPE int4 VALUE '1073741827',
        read            TYPE int4 VALUE '1073741826',
        restricted_view TYPE int4 VALUE '1073741832',
        limited_access  TYPE int4 VALUE '1073741825',
      END OF cs_id_roles .

    METHODS constructor
      IMPORTING
        !iv_domain              TYPE any
        !iv_site                TYPE any
        !iv_client_id           TYPE any
        !iv_client_secret       TYPE any
        !iv_protocol_https      TYPE sap_bool DEFAULT abap_true
        !iv_realm               TYPE any
        !iv_document_library    TYPE any OPTIONAL
        !iv_generate_tokeh_auth TYPE sap_bool DEFAULT abap_false
      RAISING
        zcx_ca_http_services .
    METHODS generate_token_auth
      RETURNING
        VALUE(rv_token) TYPE string
      RAISING
        zcx_ca_http_services .
    METHODS generate_digest_value
      RETURNING
        VALUE(rv_digest_value) TYPE string
      RAISING
        zcx_ca_http_services .
    METHODS create_document_library
      IMPORTING
        !iv_name                   TYPE any
        !iv_break_role_inheritance TYPE sap_bool DEFAULT abap_true
      EXPORTING
        !ev_doc_library_created    TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS get_doc_library_info
      IMPORTING
        !iv_name TYPE any
      EXPORTING
        !es_info TYPE ts_receive_get_info_doc_lib
      RAISING
        zcx_ca_http_services .
    METHODS create_folder
      IMPORTING
        !iv_name                   TYPE any
        !iv_break_role_inheritance TYPE sap_bool DEFAULT abap_true
      EXPORTING
        !ev_folder_created         TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS add_file
      IMPORTING
        !iv_path_folder  TYPE any
        !iv_content_file TYPE xstring
        !iv_file_name    TYPE string
      EXPORTING
        !ev_file_upload  TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS get_folders
      IMPORTING
        !iv_path_folder TYPE any OPTIONAL
      EXPORTING
        !et_folders     TYPE tt_receive_get_folders_values
      RAISING
        zcx_ca_http_services .
    METHODS get_files
      IMPORTING
        !iv_path_folder TYPE any OPTIONAL
      EXPORTING
        !et_files       TYPE tt_receive_get_files_values
      RAISING
        zcx_ca_http_services .
    METHODS delete_file
      IMPORTING
        !iv_path_folder   TYPE any OPTIONAL
        !iv_file_name     TYPE string
        !iv_verify_delete TYPE sap_bool DEFAULT abap_false
      EXPORTING
        !ev_file_deleted  TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS delete_folder
      IMPORTING
        !iv_path_folder    TYPE any OPTIONAL
        !iv_folder_name    TYPE any
        !iv_verify_delete  TYPE sap_bool DEFAULT abap_false
      EXPORTING
        !ev_folder_deleted TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS break_role_inheritance
      IMPORTING
        !iv_path_folder TYPE any
      RAISING
        zcx_ca_http_services .
    METHODS break_role_inh_doc_library
      IMPORTING
        !iv_doc_library TYPE any
      RAISING
        zcx_ca_http_services .
    METHODS get_user_info
      IMPORTING
        !iv_user_cloud   TYPE string OPTIONAL
        !iv_user_windows TYPE string OPTIONAL
        !iv_user_saml    TYPE string OPTIONAL
        !iv_user_id      TYPE int4 OPTIONAL
      EXPORTING
        !es_user_info    TYPE ts_receive_get_user_info
      RAISING
        zcx_ca_http_services .
    METHODS add_user_folder
      IMPORTING
        !iv_user_cloud   TYPE string OPTIONAL
        !iv_user_windows TYPE string OPTIONAL
        !iv_user_saml    TYPE string OPTIONAL
        !iv_user_id      TYPE int4 OPTIONAL
        !iv_role         TYPE int4 DEFAULT cs_id_roles-contribute
        !iv_path_folder  TYPE string
      EXPORTING
        !ev_role_added   TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS add_user_library
      IMPORTING
        !iv_user_cloud   TYPE string OPTIONAL
        !iv_user_windows TYPE string OPTIONAL
        !iv_user_saml    TYPE string OPTIONAL
        !iv_user_id      TYPE int4 OPTIONAL
        !iv_role         TYPE int4 DEFAULT cs_id_roles-contribute
        !iv_library      TYPE string
      EXPORTING
        !ev_role_added   TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS add_group_folder
      IMPORTING
        !iv_group       TYPE string OPTIONAL
        !iv_group_id    TYPE int4 OPTIONAL
        !iv_role        TYPE int4 DEFAULT cs_id_roles-contribute
        !iv_path_folder TYPE string
      EXPORTING
        !ev_role_added  TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS add_group_library
      IMPORTING
        !iv_group      TYPE string OPTIONAL
        !iv_group_id   TYPE int4 OPTIONAL
        !iv_role       TYPE int4 DEFAULT cs_id_roles-contribute
        !iv_library    TYPE string
      EXPORTING
        !ev_role_added TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS remove_user_folder
      IMPORTING
        !iv_user_cloud   TYPE string OPTIONAL
        !iv_user_windows TYPE string OPTIONAL
        !iv_user_saml    TYPE string OPTIONAL
        !iv_id_user      TYPE int4 OPTIONAL
        !iv_path_folder  TYPE string
      EXPORTING
        !ev_user_removed TYPE sap_bool.
    METHODS remove_user_library
      IMPORTING
        !iv_user_cloud   TYPE string OPTIONAL
        !iv_user_windows TYPE string OPTIONAL
        !iv_user_saml    TYPE string OPTIONAL
        !iv_id_user      TYPE int4 OPTIONAL
        !iv_library      TYPE string
      EXPORTING
        !ev_user_removed TYPE sap_bool
      RAISING
        zcx_ca_http_services .
    METHODS get_user_folder
      IMPORTING
        !iv_path_folder   TYPE string
        !iv_get_role      TYPE sap_bool DEFAULT abap_true
        !iv_get_loginname TYPE sap_bool DEFAULT abap_true
      EXPORTING
        !et_users         TYPE tt_get_user_folder
      RAISING
        zcx_ca_http_services .
    METHODS get_user_library
      IMPORTING
        !iv_library       TYPE string
        !iv_get_role      TYPE sap_bool DEFAULT abap_true
        !iv_get_loginname TYPE sap_bool DEFAULT abap_true
      EXPORTING
        !et_users         TYPE tt_get_user_folder
      RAISING
        zcx_ca_http_services .
    METHODS get_user_folder_role
      IMPORTING
        !iv_path_folder TYPE string
        !iv_user_id     TYPE int4
      EXPORTING
        !et_roles       TYPE tt_get_user_folder_role.
    METHODS get_user_library_role
      IMPORTING
        !iv_library TYPE string
        !iv_user_id TYPE int4
      EXPORTING
        !et_roles   TYPE tt_get_user_folder_role
      RAISING
        zcx_ca_http_services .
    METHODS get_user_folder_member
      IMPORTING
        !iv_path_folder TYPE string
        !iv_user_id     TYPE int4
      EXPORTING
        !es_member      TYPE ts_get_user_folder_member.
    METHODS get_user_library_member
      IMPORTING
        !iv_library TYPE string
        !iv_user_id TYPE int4
      EXPORTING
        !es_member  TYPE ts_get_user_folder_member
      RAISING
        zcx_ca_http_services .
    METHODS get_site_groups
      EXPORTING
        !et_groups TYPE tt_recieve_sitegroup_id
      RAISING
        zcx_ca_http_services .
    METHODS get_site_groups_by_name
      IMPORTING
        !iv_name  TYPE string
      EXPORTING
        !es_group TYPE ts_recieve_sitegroup_id
      RAISING
        zcx_ca_http_services .
    CLASS-METHODS replace_strange_char
      IMPORTING
        !iv_input        TYPE any
      RETURNING
        VALUE(rv_output) TYPE string .
    METHODS set_document_library
      IMPORTING iv_name TYPE any.
  PROTECTED SECTION.

    TYPES:
      BEGIN OF ts_receive_digest_value,
        form_digest_timeout_seconds TYPE i,
        form_digest_value           TYPE string,
        library_version             TYPE string,
        site_full_url               TYPE i,
        web_full_url                TYPE string,
      END OF ts_receive_digest_value .
    TYPES:
      BEGIN OF ts_send_create_folder,
        BEGIN OF metadata,
          type TYPE string,
        END OF metadata,
        _server_relative_url TYPE string,
      END OF ts_send_create_folder .
    TYPES:
      BEGIN OF ts_receive_create_folder,
        exists            TYPE string,
        is_wopi_enabled   TYPE string,
        itemcount         TYPE i,
        name              TYPE string,
        progid            TYPE string,
        serverrelativeurl TYPE string,
        timecreated       TYPE string,
        timelastmodified  TYPE string,
        uniqueid          TYPE string,
        welcomepage       TYPE string,
      END OF ts_receive_create_folder .
    TYPES: BEGIN OF ts_send_create_doc_library,
             BEGIN OF metadata,
               type TYPE string,
             END OF metadata,
             _allow_content_types   TYPE string,
             _base_template         TYPE i,
             _content_types_enabled TYPE string,
             _description           TYPE string,
             _title                 TYPE string,
           END OF ts_send_create_doc_library .
    TYPES:
      BEGIN OF ts_receive_create_doc_library,
        allow_content_types     TYPE string,
        base_template           TYPE i,
        base_type               TYPE i,
        content_types_enabled   TYPE string,
        crawl_non_default_views TYPE string,
        created                 TYPE string,
        timecreated             TYPE string,
        current_change_token    TYPE string,
        disablegridediting      TYPE string,
        document_template_url   TYPE string,
        id                      TYPE string,
        title                   TYPE string,
      END OF ts_receive_create_doc_library .
    TYPES:
      BEGIN OF ts_receive_add_file,
        check_in_comment       TYPE string,
        check_out_type         TYPE i,
        content_tag            TYPE string,
        customized_page_status TYPE i,
        e_tag                  TYPE string,
        exists                 TYPE string,
        irm_enabled            TYPE string,
        length                 TYPE string,
        level                  TYPE i,
        linking_uri            TYPE string,
        linking_url            TYPE string,
        major_version          TYPE i,
        minor_version          TYPE i,
        name                   TYPE string,
        server_relative_url    TYPE string,
        time_created           TYPE string,
        time_last_modified     TYPE string,
        title                  TYPE string,
        ui_version             TYPE i,
        ui_version_label       TYPE string,
        unique_id              TYPE string,
      END OF ts_receive_add_file .
    TYPES:
      BEGIN OF ts_receive_get_folders,
        value TYPE tt_receive_get_folders_values,
      END OF ts_receive_get_folders .
    TYPES:
      BEGIN OF ts_receive_get_files,
        value TYPE tt_receive_get_files_values,
      END OF ts_receive_get_files .
    TYPES:
      BEGIN OF ts_receive_get_user_folder_id,
        principal_id TYPE int4,
      END OF ts_receive_get_user_folder_id .
    TYPES:
      tt_receive_get_user_folder_id TYPE STANDARD TABLE OF ts_receive_get_user_folder_id WITH EMPTY KEY .
    TYPES:
      BEGIN OF ts_receive_get_user_folder,
        value TYPE tt_receive_get_user_folder_id,
      END OF ts_receive_get_user_folder .
    TYPES:
      BEGIN OF ts_recieve_sitegroup,
        value TYPE tt_recieve_sitegroup_id,
      END OF ts_recieve_sitegroup .
    TYPES:
      BEGIN OF ts_receive_get_usrfold_role,
        value TYPE tt_get_user_folder_role,
      END OF ts_receive_get_usrfold_role .

    DATA mv_token_auth TYPE string .
    DATA mv_digest_value TYPE string .
    DATA mv_client_id TYPE string .
    DATA mv_client_secret TYPE string .
    DATA mv_realm TYPE string .
    DATA mv_domain TYPE string .
    DATA mv_http_protocol TYPE string .
    DATA mv_site TYPE string .
    DATA mv_document_library TYPE string .

    METHODS verify_generate_tokens
      IMPORTING
        !iv_token_auth   TYPE sap_bool DEFAULT abap_true
        !iv_token_digest TYPE sap_bool DEFAULT abap_true
      RAISING
        zcx_ca_http_services .
    METHODS delete_slash_begin_end
      CHANGING
        !cv_path TYPE string .
    METHODS escape_url
      IMPORTING
        !iv_input        TYPE any
      RETURNING
        VALUE(rv_output) TYPE string .
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_ca_sharepoint_rest IMPLEMENTATION.


  METHOD add_file.
    DATA ls_receive TYPE ts_receive_add_file.

    ev_file_upload = abap_true.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    DATA(lv_content_len) = xstrlen( iv_content_file ). " Se calcula el contenido del binario pasado

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
    set_header_value( iv_name  = 'Content-Type' iv_value = 'application/json;odata=verbose' ).
    set_header_value( iv_name  = 'content-length' iv_value = lv_content_len ).

    " Se pasa los datos a la petición
    set_data( EXPORTING iv_data   = iv_content_file
                        iv_length = lv_content_len ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica a que URL se va a llamar
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/Files/add(url='{ escape_url( iv_file_name ) }',overwrite=true)| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Si llega hasta aquí producirse ninguna excepción es que con toda seguridad se haya creado. Aún así, miro si
    " un campo de retorno que indica si existe.
    ev_file_upload = ls_receive-exists.

  ENDMETHOD.


  METHOD add_group_folder.
    DATA ls_group_info TYPE ts_recieve_sitegroup_id.

    " Si el id del grupo se pasa no es necesario buscarlo, pero si se pasa el nombre se busca para saber su ID
    IF iv_group_id IS NOT INITIAL.
      ls_group_info-id = iv_group_id.
    ELSE.
      get_site_groups_by_name(
        EXPORTING
          iv_name  = iv_group
        IMPORTING
          es_group = ls_group_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).

    " El montaje del usuario y rol
    DATA(lv_param_role) = |principalid={ ls_group_info-id },roleDefId={ iv_role }|.
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments/addroleassignment({ lv_param_role })| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " El resultado del servicio no es nada descriptivo. Devuelve esto: 'odata.null': true. Como muchos servicios
    " de sharepoint las respuestas no son muy claras. Y aquí menos.
    receive( ).

    " Si llega aquí es que el rol se ha asignado porque no se ha generado excepción
    ev_role_added = abap_true.
  ENDMETHOD.


  METHOD add_user_folder.
    DATA ls_user_info TYPE ts_receive_get_user_info.

* NOTA: Para poder asignar un usuario a una carpeta previamente se tiene que haber parado la herencia de roles sino no funcionara.
* para eso esta el método BREAK_ROLE_INHERITANCE que se encarga de eso. Esa acción no se hace aquí porque el objetivo del método
* es ir añadiendo usuario a una carpeta, de una manera simple.


    ev_role_added = abap_false.

* Si se ha pasado el ID del usuario se usa, en caso contrario se obtiene del usuario login pasado
    IF iv_user_id IS NOT INITIAL.
      ls_user_info-id = iv_user_id.
    ELSE.
      get_user_info( EXPORTING iv_user_cloud   = iv_user_cloud
                               iv_user_windows =  iv_user_windows
                               iv_user_saml    = iv_user_saml
                     IMPORTING es_user_info    = ls_user_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ lv_path_folder }| ).

    " El montaje del usuario y rol
    DATA(lv_param_role) = |principalid={ ls_user_info-id },roleDefId={ iv_role }|.

    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments/addroleassignment({ lv_param_role })| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " El resultado del servicio no es nada descriptivo. Devuelve esto: 'odata.null': true. Como muchos servicios
    " de sharepoint las respuestas no son muy claras. Y aquí menos.
    receive( ).

    " Si llega aquí es que el rol se ha asignado porque no se ha generado excepción
    ev_role_added = abap_true.
  ENDMETHOD.


  METHOD break_role_inheritance.
    DATA lv_receive TYPE string.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).


    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
*    set_header_value( iv_name  = 'Content-Type' iv_value = 'application/json;odata=verbose' ).
    set_header_value( iv_name  = 'content-length' iv_value = 0 ).

    " El la librería se escapa porque se pasa como URL
    DATA(lv_path) = escape_url( mv_document_library ).

    " la ruta donde esta el directorio a comprobar se le quita el carácter / inicial y final porque se añadirá en este método
    IF iv_path_folder IS NOT INITIAL.
      DATA(lv_path_folder) = CONV string( iv_path_folder ).
      delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

      " Se añade la libreria al path pasado de lectura. Se escapa todo para evitar problemas
      TRY.
          lv_path = |{ lv_path }/{ escape_url( lv_path_folder ) }|.
        CATCH cx_root.
          lv_path = |{ lv_path }/{ lv_path_folder }|.
      ENDTRY.

    ENDIF.

    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/breakroleinheritance(copyroleassignments=false)| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " El resultado del servicio no es nada descriptivo. Devuelve esto: 'odata.null': true. Como muchos servicios
    " de sharepoint las respuestas son muy claras. Y aquí menos.
    receive( IMPORTING ev_data = lv_receive ).

  ENDMETHOD.


  METHOD constructor.
    super->constructor( ).

    " Campos fijos
    mv_client_id = iv_client_id.
    mv_realm = iv_realm. " Es el id unico asignado al servidor online de sharepoint
    mv_site = iv_site.

    mv_document_library = iv_document_library.
    REPLACE ALL OCCURRENCES OF '/' IN mv_document_library WITH space. " Se quitan las / porque afectarán a las llamadas
    CONDENSE mv_document_library.

    mv_client_secret = iv_client_secret.

    " Protolo http dependedel tipo de conexión
    mv_http_protocol = COND #( WHEN iv_protocol_https = abap_true THEN |https://| ELSE |http://| ).

    " Al dominio se le quita los posibles http o https que pueda tener
    mv_domain = |{ iv_domain CASE = LOWER }|.
    REPLACE ALL OCCURRENCES OF 'http://' IN mv_domain WITH space.
    REPLACE ALL OCCURRENCES OF 'https://' IN mv_domain WITH space.
    CONDENSE mv_domain.

    " Se genera la clase que gestionará las peticiones http
    create_http_client( iv_host = mv_domain iv_is_https = iv_protocol_https ).

    " Se genera el token de autentificación si se especifica por parámetro
    IF iv_generate_tokeh_auth = abap_true.
      mv_token_auth = generate_token_auth( ).
    ENDIF.

  ENDMETHOD.


  METHOD create_document_library.
    DATA ls_receive TYPE ts_receive_create_doc_library.
    DATA ls_send TYPE ts_send_create_doc_library.

    ev_doc_library_created = abap_false.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Parámetro para la creación de la document library
    ls_send-metadata-type = 'SP.List'.
    ls_send-_allow_content_types = |true|.
    ls_send-_content_types_enabled = |true|.
    ls_send-_base_template = cs_base_template-document_library.
    ls_send-_description = ls_send-_title = iv_name.

    " Se convierte la estructura a JSON. El motivo es que hay un tag que no puede se puede definir con la clase de SAP. Por lo tanto luego
    " hago un replace para hacerlo compatible.
    DATA(lv_json_send) = /ui2/cl_json=>serialize( data = ls_send compress = abap_true pretty_name = /ui2/cl_json=>pretty_mode-camel_case ).
    REPLACE FIRST OCCURRENCE OF 'metadata' IN lv_json_send WITH '__metadata'.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
    set_header_value( iv_name  = 'Content-Type' iv_value = 'application/json;odata=verbose' ).
    set_header_value( iv_name  = 'content-length' iv_value = strlen( lv_json_send ) ).

    " Se indica a que URL se va a llamar
    set_request_uri( |/sites/{ mv_site }/_api/web/lists| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( iv_data = lv_json_send ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Si llega hasta aquí producirse ninguna excepción es que con toda seguridad se haya creado.
    ev_doc_library_created = abap_true.

    " Pongo en la variable global la libreria creada
    set_document_library( iv_name ).

    " Si la carpeta se crea correctament y por parámetro se quiere que la carpeta no herede los permisos del padre, entonces se llama
    " al método encargado de hacerlo.
    " Cualquier excepción se captura pero no se trata porque la prioridad es la creación de la carpeta, y devolviendo la excepción se
    " desvirtua el servicio. Además, el servicio de roles es tan tonto que dudo que falle existiendo la carpeta.
    IF ev_doc_library_created = abap_true AND iv_break_role_inheritance = abap_true.
      TRY.
          break_role_inh_doc_library( mv_document_library ).
        CATCH zcx_ca_http_services. " CA - Excepciones clase de servicios HTTP
      ENDTRY.
    ENDIF.

  ENDMETHOD.


  METHOD create_folder.
    DATA ls_receive TYPE ts_receive_create_folder.
    DATA ls_send TYPE ts_send_create_folder.

    ev_folder_created = abap_false.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido.
    DATA(lv_path_folder) = CONV string( iv_name ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Parámetro para la creación de la carpeta
    ls_send-metadata-type = 'SP.Folder'.
    ls_send-_server_relative_url = escape_url( |{ mv_http_protocol }{ mv_domain }/sites/{ mv_site }/{ mv_document_library }/{ lv_path_folder }| ).

    " Se convierte la estructura a JSON. El motivo es que hay un tag que no puede se puede definir con la clase de SAP. Por lo tanto luego
    " hago un replace para hacerlo compatible.
    DATA(lv_json_send) = /ui2/cl_json=>serialize( data = ls_send compress = abap_true pretty_name = /ui2/cl_json=>pretty_mode-camel_case ).
    REPLACE FIRST OCCURRENCE OF 'metadata' IN lv_json_send WITH '__metadata'.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
    set_header_value( iv_name  = 'Content-Type' iv_value = 'application/json;odata=verbose' ).
    set_header_value( iv_name  = 'content-length' iv_value = strlen( lv_json_send ) ).

    " Se indica a que URL se va a llamar
    set_request_uri( |/sites/{ mv_site }/_api/web/folders| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( iv_data = lv_json_send ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Si llega hasta aquí producirse ninguna excepción es que con toda seguridad se haya creado. Aún así, miro si
    " un campo de retorno que indica si existe.
    ev_folder_created = ls_receive-exists.

    " Si la carpeta se crea correctament y por parámetro se quiere que la carpeta no herede los permisos del padre, entonces se llama
    " al método encargado de hacerlo.
    " Cualquier excepción se captura pero no se trata porque la prioridad es la creación de la carpeta, y devolviendo la excepción se
    " desvirtua el servicio. Además, el servicio de roles es tan tonto que dudo que falle existiendo la carpeta.
    IF ev_folder_created = abap_true AND iv_break_role_inheritance = abap_true.
      TRY.
          break_role_inheritance( lv_path_folder ).
        CATCH zcx_ca_http_services. " CA - Excepciones clase de servicios HTTP
      ENDTRY.
    ENDIF.

  ENDMETHOD.


  METHOD delete_file.

    ev_file_deleted = abap_false.

    " Tanto el site como la libreria sea escapa
    DATA(lv_path) = escape_url( |/sites/{ mv_site }/{ mv_document_library }| ).

    " la ruta donde esta el directorio a comprobar se le quita el carácter / inicial y final porque se añadirá en este método
    IF iv_path_folder IS NOT INITIAL.
      DATA(lv_path_folder) = CONV string( iv_path_folder ).
      delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

      " Se añade el path al path de lectura
      " En determinados casos el escapar ciertos valore se produce una excepción, en ese caso se pondra el mismo valor
      TRY.
          lv_path = |{ lv_path }/{ escape_url( lv_path_folder ) }|.
        CATCH cx_root.
          lv_path = |{ lv_path }/{ lv_path_folder }|.
      ENDTRY.


    ENDIF.


    " Se informa la consulta a realizar
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFileByServerRelativeUrl('{ lv_path }/{ cl_http_utility=>escape_url( iv_file_name ) }')| ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
    set_header_value( iv_name  = 'IF-MATCH' iv_value = '*' ).
    set_header_value( iv_name  = 'X-HTTP-Method' iv_value = 'DELETE' ).
    set_header_value( iv_name  = 'content-length' iv_value = 0 ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, que en el borrado solo se recupera cualquier posible excepcion. Pero en las pruebas se ha visto que no
    " peta ni aunque le pongas un directorio o fichero que no existe.
    receive( ).

    " Por eso con el siguiente parámetro se hace la verificación que el fichero exista. La verficiación consiste en leer
    " los archivos y chequear que exista. Si no existe, se ha borrado.
    IF iv_verify_delete = abap_true.

      get_files( EXPORTING iv_path_folder = iv_path_folder
                 IMPORTING et_files =  DATA(lt_files) ).

      READ TABLE lt_files TRANSPORTING NO FIELDS WITH KEY name = iv_file_name.
      IF sy-subrc NE 0.
        ev_file_deleted = abap_true.
      ELSE.
        ev_file_deleted = abap_false.
      ENDIF.

    ELSE. " Si no se quiere hacer la verificación se devuelve que se ha borrado
      ev_file_deleted = abap_true.
    ENDIF.

  ENDMETHOD.


  METHOD delete_folder.

    ev_folder_deleted = abap_false.

    " Tanto el site como la libreria sea escapa
    DATA(lv_path) = escape_url( |/sites/{ mv_site }/{ mv_document_library }| ).

    " la ruta donde esta el directorio a comprobar se le quita el carácter / inicial y final porque se añadirá en este método
    IF iv_path_folder IS NOT INITIAL.
      DATA(lv_path_folder) = CONV string( iv_path_folder ).
      delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

      " Se añade el path al path de lectura
      lv_path = |{ lv_path }/{ cl_http_utility=>escape_url( lv_path_folder ) }|.
    ENDIF.

    " En determinados casos el escapar ciertos valore se produce una excepción, en ese caso se pondra el mismo valor
    TRY.
        DATA(lv_folder) = escape_url( iv_folder_name ).
      CATCH cx_root.
        lv_folder = iv_folder_name.
    ENDTRY.

    " Se informa la consulta a realizar
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }/{ lv_folder }')| ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
    set_header_value( iv_name  = 'IF-MATCH' iv_value = '*' ).
    set_header_value( iv_name  = 'X-HTTP-Method' iv_value = 'DELETE' ).
    set_header_value( iv_name  = 'content-length' iv_value = 0 ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, que en el borrado solo se recupera cualquier posible excepcion. Pero en las pruebas se ha visto que no
    " peta ni aunque le pongas un directorio o fichero que no existe.
    receive( ).

    " Por eso con el siguiente parámetro se hace la verificación que el fichero exista. La verficiación consiste en leer
    " los archivos y chequear que exista. Si no existe, se ha borrado.
    IF iv_verify_delete = abap_true.

      get_folders( EXPORTING iv_path_folder = iv_path_folder
                   IMPORTING et_folders =  DATA(lt_folders) ).

      READ TABLE lt_folders TRANSPORTING NO FIELDS WITH KEY name = iv_folder_name.


      IF sy-subrc NE 0.
        ev_folder_deleted = abap_true.
      ELSE.
        ev_folder_deleted = abap_false.
      ENDIF.

    ELSE. " Si no se quiere hacer la verificación se devuelve que se ha borrado
      ev_folder_deleted = abap_true.
    ENDIF.
  ENDMETHOD.


  METHOD delete_slash_begin_end.
    IF cv_path(1) = '/'. " El '/' del principio se elimina
      REPLACE SECTION OFFSET 0 LENGTH 1 OF  cv_path  WITH space.
    ENDIF.

    " Calculo la longitud del path para poder mirar si el último dígito tiene el '/'
    DATA(lv_len) = strlen( cv_path  ) - 1.
    IF cv_path+lv_len(1) = '/'.
      REPLACE SECTION OFFSET lv_len LENGTH 1 OF cv_path WITH space.
    ENDIF.
  ENDMETHOD.


  METHOD escape_url.

    rv_output = cl_http_utility=>escape_url( replace_strange_char( iv_input ) ).

  ENDMETHOD.


  METHOD generate_digest_value.

    DATA ls_receive_digest_value TYPE ts_receive_digest_value.

    CLEAR rv_digest_value.

    " Para generar el token digest hace falta el de autentificación. Por lo que llama al método que se encarga
    " de verificar y de crear en dicho caso el token. Como el método hace lo mismo con el propio digest, ya que se usa
    " en varios sitios, le tengo que indicar que no lo valide porque haría un bucle recursivo.
    verify_generate_tokens( iv_token_digest = abap_false ).

    " Datos de cabecera para obtener el digest
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'content-length' iv_value = 0 ).


    " Se indica a que URL se va a llamar
    set_request_uri( |/sites/{ mv_site }/_api/contextinfo| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive_digest_value ).

    mv_digest_value = cl_http_utility=>escape_url( ls_receive_digest_value-form_digest_value ).

    " Lo que se devuelve no se escapa para quien lo reciba lo trate como le interesa.
    rv_digest_value = ls_receive_digest_value-form_digest_value.

  ENDMETHOD.


  METHOD generate_token_auth.

    CLEAR rv_token.

    rv_token = NEW zcl_ca_microsoft_auth( )->generate_token_auth( iv_client_secret = mv_client_secret
                                   iv_client_id = mv_client_id
                                   iv_domain = mv_domain
                                   iv_realm = mv_realm ).

  ENDMETHOD.


  METHOD get_files.
    DATA ls_receive TYPE ts_receive_get_files.

    CLEAR et_files.

    " Se lanza al método que verifica que esten los tokens necesarios. Como es de lectura el digest no hace falta.
    verify_generate_tokens( iv_token_digest = abap_false ).


    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " El la librería se escapa porque se pasa como URL
    DATA(lv_path) = escape_url( mv_document_library ).

    " la ruta donde esta el directorio a comprobar se le quita el carácter / inicial y final porque se añadirá en este método
    IF iv_path_folder IS NOT INITIAL.
      DATA(lv_path_folder) = CONV string( iv_path_folder ).
      delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

      TRY.
          lv_path = |{ lv_path }/{ escape_url( lv_path_folder ) }|.
        CATCH cx_root.
          lv_path = |{ lv_path }/{ lv_path_folder }|.
      ENDTRY.

    ENDIF.

    " Se informa la consulta a realizar
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/Files| ).

    " El método GET
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " El formato se pasa una tabla con la misma estructura pero sin el nodo intermedio values
    LOOP AT ls_receive-value ASSIGNING FIELD-SYMBOL(<ls_value>).
      INSERT <ls_value> INTO TABLE et_files.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_folders.
    DATA ls_receive TYPE ts_receive_get_folders.

    CLEAR et_folders.

    " Se lanza al método que verifica que esten los tokens necesarios. Como es de lectura el digest no hace falta.
    verify_generate_tokens( iv_token_digest = abap_false ).


    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " El la librería se escapa porque se pasa como URL
    DATA(lv_path) = escape_url( mv_document_library ).

    " la ruta donde esta el directorio a comprobar se le quita el carácter / inicial y final porque se añadirá en este método
    IF iv_path_folder IS NOT INITIAL.
      DATA(lv_path_folder) = CONV string( iv_path_folder ).
      delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

      " Se añade la libreria al path pasado de lectura. Se escapa todo para evitar problemas
      TRY.
          lv_path = |{ lv_path }/{ escape_url( lv_path_folder ) }|.
        CATCH cx_root.
          lv_path = |{ lv_path }/{ lv_path_folder }|.
      ENDTRY.

    ENDIF.

    " Se informa la consulta a realizar
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/Folders| ).

    " El método GET
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " El formato se pasa una tabla con la misma estructura pero sin el nodo intermedio values
    LOOP AT ls_receive-value ASSIGNING FIELD-SYMBOL(<ls_value>).
      INSERT <ls_value> INTO TABLE et_folders.
    ENDLOOP.

  ENDMETHOD.


  METHOD get_site_groups.
    DATA ls_receive TYPE ts_recieve_sitegroup.

    CLEAR et_groups.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/sitegroups| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).


    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Se pasa los datos recogidos a la estructura de salida
    et_groups = ls_receive-value.
  ENDMETHOD.


  METHOD get_site_groups_by_name.
    CLEAR es_group.

    CLEAR es_group.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/sitegroups/getbyname('{ escape_url( iv_name ) }')| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = es_group ).

  ENDMETHOD.


  METHOD get_user_folder.
    DATA ls_receive TYPE ts_receive_get_user_folder.

    CLEAR et_users.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Se pasa los datos recogidos a la estructura de salida
    LOOP AT ls_receive-value ASSIGNING FIELD-SYMBOL(<ls_values>).
      APPEND INITIAL LINE TO et_users ASSIGNING FIELD-SYMBOL(<ls_user>).
      <ls_user>-user_id = <ls_values>-principal_id.

      " Se recupera los roles que tiene asignado en la carpeta
      IF iv_get_role = abap_true.
        TRY.
            get_user_folder_role(
              EXPORTING
                iv_path_folder = iv_path_folder
                iv_user_id     = <ls_user>-user_id
              IMPORTING
                et_roles       =  <ls_user>-role ).
          CATCH zcx_ca_http_services.
        ENDTRY.

      ENDIF.
      IF iv_get_loginname = abap_true.
        TRY.
            get_user_folder_member(
             EXPORTING
               iv_path_folder = iv_path_folder
               iv_user_id     = <ls_user>-user_id
             IMPORTING
               es_member = DATA(ls_member) ).
            " Si el el miembro es un usuario el campo de email estará informado, en caso contrario, será u
            " un grupo y habrá que tomar el campo "login_name".
            IF ls_member-email IS NOT INITIAL.
              <ls_user>-user_loginname = ls_member-email.
            ELSE.
              <ls_user>-user_loginname = ls_member-login_name.
            ENDIF.
          CATCH zcx_ca_http_services.
        ENDTRY.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD get_user_folder_member.


    CLEAR es_member.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments({ iv_user_id })/Member| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = es_member ).


  ENDMETHOD.


  METHOD get_user_folder_role.
    DATA ls_receive TYPE ts_receive_get_usrfold_role.

    CLEAR et_roles.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments({ iv_user_id })/RoleDefinitionBindings| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Se pasa los datos recogidos a la estructura de salida

    et_roles = ls_receive-value.
  ENDMETHOD.


  METHOD get_user_info.


    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Obtener los datos del usuario por el ID implica llamar a un servicio distinto, pero el resultado es el mismo
    " que si se llama por el loginname. Por lo tanto, se hace a la vez y es más simple las llamadas
    IF iv_user_id IS NOT INITIAL.
      set_request_uri( |/sites/{ mv_site }/_api/web/getuserbyid({ iv_user_id })| ).
    ELSE.
      " Se monta el login según el parámetro informado. Si no se pasa nada se sale del proceso
      IF iv_user_cloud IS NOT INITIAL.
        DATA(lv_login) = |i:0#.f\|membership\|{ iv_user_cloud }|.
      ELSEIF iv_user_windows IS NOT INITIAL.
        lv_login = |i:0#.w\|{ iv_user_windows }|.
      ELSEIF iv_user_saml IS NOT INITIAL.
        lv_login = |i:05.t\|{ iv_user_saml }|.
      ELSE.
        EXIT.
      ENDIF.

      " Llamada para obtener por el loginname
      set_request_uri( |/sites/{ mv_site }/_api/web/siteusers('{ escape_url( lv_login ) }')| ).

    ENDIF.

    " El método GET
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = es_user_info ).

  ENDMETHOD.


  METHOD remove_user_folder.
    DATA ls_user_info TYPE ts_receive_get_user_info.

* NOTA: Para poder asignar un usuario a una carpeta previamente se tiene que haber parado la herencia de roles sino no funcionara.
* para eso esta el método BREAK_ROLE_INHERITANCE que se encarga de eso. Esa acción no se hace aquí porque el objetivo del método
* es ir añadiendo usuario a una carpeta, de una manera simple.


    ev_user_removed = abap_false.

* Si se ha pasado el ID del usuario se usa, en caso contrario se obtiene del usuario login pasado
    IF iv_id_user IS NOT INITIAL.
      ls_user_info-id = iv_id_user.

      " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
      verify_generate_tokens( ).

    ELSE.
      get_user_info( EXPORTING iv_user_cloud   = iv_user_cloud
                               iv_user_windows =  iv_user_windows
                               iv_user_saml    = iv_user_saml
                     IMPORTING es_user_info    = ls_user_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " Se quitan los posibles carácteres iniciales y finales que pueda tener el path introducido. En este servicio
    " se ha comprobado que la barra oblicua(/) en el path no hace que falle. Pero se quita "por si acaso".
    DATA(lv_path_folder) = CONV string( iv_path_folder ).
    delete_slash_begin_end( CHANGING cv_path = lv_path_folder ).

    " Se indica el path
    DATA(lv_path) = escape_url( |{ mv_document_library }/{ cl_http_utility=>escape_url( lv_path_folder ) }| ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/GetFolderByServerRelativeUrl('{ lv_path }')/ListItemAllFields/roleassignments/getbyprincipalid({ ls_user_info-id })| ).

    " El método POST
    set_request_method( 'DELETE' ).

    " Se hace el envio
    send( ).

    " No hay resultado del servicio, no devuelve nada.
    receive( ).

    " Si llega aquí es que el rol se ha desasignado porque no se ha generado excepción
    ev_user_removed = abap_true.
  ENDMETHOD.


  METHOD replace_strange_char.

    CLEAR rv_output.

    rv_output = CONV string( iv_input ).

    " Se quitán caracteres extraños que afectan al sharepoint
    REPLACE ALL OCCURRENCES OF |'| IN rv_output WITH '' .
    REPLACE ALL OCCURRENCES OF |*| IN rv_output WITH '' .


  ENDMETHOD.


  METHOD verify_generate_tokens.

    IF iv_token_auth = abap_true AND mv_token_auth IS INITIAL.
      mv_token_auth = generate_token_auth( ).
    ENDIF.

    IF iv_token_digest = abap_true AND mv_digest_value IS INITIAL.
      generate_digest_value( ).
    ENDIF.

  ENDMETHOD.
  METHOD get_doc_library_info.
    DATA ls_receive TYPE ts_receive_create_doc_library.

    CLEAR: es_info.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_name ) }')| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = es_info ).


  ENDMETHOD.

  METHOD set_document_library.
    mv_document_library = iv_name.
  ENDMETHOD.

  METHOD break_role_inh_doc_library.


    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).
    set_header_value( iv_name  = 'X-RequestDigest' iv_value = mv_digest_value ).
*    set_header_value( iv_name  = 'Content-Type' iv_value = 'application/json;odata=verbose' ).
    set_header_value( iv_name  = 'content-length' iv_value = 0 ).

    " URL
    DATA(lv_url) = |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_doc_library ) }')/breakroleinheritance(copyroleassignments=false)|.
    set_request_uri( lv_url ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " Se recupera el resultado, teniendo en cuenta que hay a nivel interno se escapa como si fuesa una URL porque se envia posteriormente
    " como cabecera http.
    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case ).
  ENDMETHOD.

  METHOD add_user_library.
    DATA ls_user_info TYPE ts_receive_get_user_info.

* NOTA: Para poder asignar un usuario a una carpeta previamente se tiene que haber parado la herencia de roles sino no funcionara.
* para eso esta el método BREAK_ROLE_INH_DOC_LIBRARY que se encarga de eso. Esa acción no se hace aquí porque el objetivo del método
* es ir añadiendo usuario a una carpeta, de una manera simple.


    ev_role_added = abap_false.

* Si se ha pasado el ID del usuario se usa, en caso contrario se obtiene del usuario login pasado
    IF iv_user_id IS NOT INITIAL.
      ls_user_info-id = iv_user_id.
    ELSE.
      get_user_info( EXPORTING iv_user_cloud   = iv_user_cloud
                               iv_user_windows =  iv_user_windows
                               iv_user_saml    = iv_user_saml
                     IMPORTING es_user_info    = ls_user_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " El montaje del usuario y rol
    DATA(lv_param_role) = |principalid={ ls_user_info-id },roleDefId={ iv_role }|.

    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments/addroleassignment({ lv_param_role })| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " El resultado del servicio no es nada descriptivo. Devuelve esto: 'odata.null': true. Como muchos servicios
    " de sharepoint las respuestas no son muy claras. Y aquí menos.
    receive( ).

    " Si llega aquí es que el rol se ha asignado porque no se ha generado excepción
    ev_role_added = abap_true.
  ENDMETHOD.

  METHOD add_group_library.
    DATA ls_group_info TYPE ts_recieve_sitegroup_id.

    " Si el id del grupo se pasa no es necesario buscarlo, pero si se pasa el nombre se busca para saber su ID
    IF iv_group_id IS NOT INITIAL.
      ls_group_info-id = iv_group_id.
    ELSE.
      get_site_groups_by_name(
        EXPORTING
          iv_name  = iv_group
        IMPORTING
          es_group = ls_group_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " El montaje del usuario y rol
    DATA(lv_param_role) = |principalid={ ls_group_info-id },roleDefId={ iv_role }|.
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments/addroleassignment({ lv_param_role })| ).

    " El método POST
    set_request_method( 'POST' ).

    " Se hace el envio
    send( ).

    " El resultado del servicio no es nada descriptivo. Devuelve esto: 'odata.null': true. Como muchos servicios
    " de sharepoint las respuestas no son muy claras. Y aquí menos.
    receive( ).

    " Si llega aquí es que el rol se ha asignado porque no se ha generado excepción
    ev_role_added = abap_true.
  ENDMETHOD.

  METHOD get_user_library.
    DATA ls_receive TYPE ts_receive_get_user_folder.

    CLEAR et_users.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).


    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Se pasa los datos recogidos a la estructura de salida
    LOOP AT ls_receive-value ASSIGNING FIELD-SYMBOL(<ls_values>).
      APPEND INITIAL LINE TO et_users ASSIGNING FIELD-SYMBOL(<ls_user>).
      <ls_user>-user_id = <ls_values>-principal_id.

      " Se recupera los roles que tiene asignado en la carpeta
      IF iv_get_role = abap_true.
        TRY.
            get_user_library_role(
              EXPORTING
                iv_library = iv_library
                iv_user_id     = <ls_user>-user_id
              IMPORTING
                et_roles       =  <ls_user>-role ).
          CATCH zcx_ca_http_services.
        ENDTRY.

      ENDIF.
      IF iv_get_loginname = abap_true.
        TRY.
            get_user_library_member(
             EXPORTING
               iv_library = iv_library
               iv_user_id     = <ls_user>-user_id
             IMPORTING
               es_member = DATA(ls_member) ).
            " Si el el miembro es un usuario el campo de email estará informado, en caso contrario, será u
            " un grupo y habrá que tomar el campo "login_name".
            IF ls_member-email IS NOT INITIAL.
              <ls_user>-user_loginname = ls_member-email.
            ELSE.
              <ls_user>-user_loginname = ls_member-login_name.
            ENDIF.
            " Tipo de usuario: 1 es usuario y 8 es grupo. Importante para saber como operar a posteriori
            <ls_user>-user_type = ls_member-principal_type.
          CATCH zcx_ca_http_services.
        ENDTRY.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD get_user_library_role.
    DATA ls_receive TYPE ts_receive_get_usrfold_role.

    CLEAR et_roles.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).

    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments({ iv_user_id })/RoleDefinitionBindings| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = ls_receive ).

    " Se pasa los datos recogidos a la estructura de salida

    et_roles = ls_receive-value.
  ENDMETHOD.

  METHOD get_user_library_member.
    CLEAR es_member.

    " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
    verify_generate_tokens( ).

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).


    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments({ iv_user_id })/Member| ).

    " El método POST
    set_request_method( 'GET' ).

    " Se hace el envio
    send( ).

    receive( EXPORTING iv_pretty_name = /ui2/cl_json=>pretty_mode-camel_case IMPORTING ev_data = es_member ).
  ENDMETHOD.

  METHOD remove_user_library.
  DATA ls_user_info TYPE ts_receive_get_user_info.

* NOTA: Para poder asignar un usuario a una carpeta previamente se tiene que haber parado la herencia de roles sino no funcionara.
* para eso esta el método BREAK_ROLE_INHERITANCE que se encarga de eso. Esa acción no se hace aquí porque el objetivo del método
* es ir añadiendo usuario a una carpeta, de una manera simple.


    ev_user_removed = abap_false.

* Si se ha pasado el ID del usuario se usa, en caso contrario se obtiene del usuario login pasado
    IF iv_id_user IS NOT INITIAL.
      ls_user_info-id = iv_id_user.

      " Se lanza al método que verifica que esten los tokens necesarios, en caso de no estarlo se llamarán para su creación
      verify_generate_tokens( ).

    ELSE.
      get_user_info( EXPORTING iv_user_cloud   = iv_user_cloud
                               iv_user_windows =  iv_user_windows
                               iv_user_saml    = iv_user_saml
                     IMPORTING es_user_info    = ls_user_info ).
    ENDIF.

    " Datos de cabecera
    set_header_value( iv_name  = 'Authorization' iv_value = mv_token_auth ).
    set_header_value( iv_name  = 'Accept' iv_value = 'application/json;odata=nometadata' ).


    " URL
    set_request_uri( |/sites/{ mv_site }/_api/web/lists/getbytitle('{ escape_url( iv_library ) }')/roleassignments/getbyprincipalid({ ls_user_info-id })| ).

    " El método POST
    set_request_method( 'DELETE' ).

    " Se hace el envio
    send( ).

    " No hay resultado del servicio, no devuelve nada.
    receive( ).

    " Si llega aquí es que el rol se ha desasignado porque no se ha generado excepción
    ev_user_removed = abap_true.
  ENDMETHOD.

ENDCLASS.
