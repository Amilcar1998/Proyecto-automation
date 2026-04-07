*** Settings ***
Library    SeleniumLibrary
Library    ../Page/Page_Login.py
Library    ../Page/page_menu.py
Library    ../OCEANO/OCEANOMAIN.py
Library    ../OCEANO/POHEADER.py

*** Variables ***
${URL}        http://was7tr1.siman.com/AccesoSUMMER/
${USUARIO}    ELOPEZ
${CLAVE}      Amilcar2025*

*** Test Cases ***
Login y Crear PO
    Open Browser    ${URL}    chrome
    Maximize Browser Window
    Login    ${USUARIO}    ${CLAVE}
    Sleep    2s
    Acceder Frame Menu
    Hacer Click En Summer
    Click Oceano
    Acceder Frame Menu
    Crear PO Menu
    Crear PO
    Acceder Frame Trabajo
    Main Descripcion
    Digitar Header PO
    Sleep    10s
    #[Teardown]    Close Browser
