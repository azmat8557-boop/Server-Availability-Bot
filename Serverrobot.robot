*** Settings ***
Library           SeleniumLibrary
Library           OperatingSystem
Library           serverpython.py    # Robot Framework  calling Python functions directly.

*** Variables ***
${URL}            https://botsdna.com/ServerAvailability/
${BROWSER}        Chrome
${MAX_EMAIL_ROWS}    5

*** Keywords ***
Open BotsDNA And Configure Downloads
    # Starts a browser session, opens the site, and sets download folder for `input.xlsx`.
    ${prefs}=      Create Dictionary    download.default_directory=${EXECDIR}
    ${options}=    Evaluate             sys.modules['selenium.webdriver'].ChromeOptions()    sys
    Call Method    ${options}           add_experimental_option    prefs    ${prefs}
    
    Create Webdriver    Chrome    options=${options}
    Go To    ${URL}
    Maximize Browser Window
    Wait Until Page Contains    Server Availability

Download The Excel Input File
    # Removes any old `input.xlsx` and downloads a fresh copy from the page.
    Remove File    ${EXECDIR}/input*.xlsx
    Click Element    xpath=//a[contains(@href, 'input.xlsx')]
    Sleep    3s

Process All Servers From Python
    # Get test rows from Python (reads `input.xlsx` and prepares values Robot needs).
    @{SERVER_DATA}=    Get Server Data

    # Process each row (one server at a time).
    FOR    ${row}    IN    @{SERVER_DATA}
        # Extract required fields from the current row.
        ${uid}=             Set Variable    ${row}[UID]
        ${pwd}=             Set Variable    ${row}[PWD]
        ${server_code}=     Set Variable    ${row}[Server Code]
        ${ip}=              Set Variable    ${row}[IP]
        
        ${server_option}=   Set Variable    ${row}[ServerOption]
        
        Input Text    id=username    ${uid}
        Input Text    id=password    ${pwd}
        
        # Some server codes might not exist on the page dropdown.
        # If selection fails, we record it and move to the next row.
        ${result}    ${msg}=    Run Keyword And Ignore Error
        ...    Select From List By Label    id=name    ${server_option}

        # If server is not available in the dropdown, mark it and skip immediately.
        Run Keyword If    '${result}' == 'FAIL'
        ...    Save Server Status    ${server_code}    NOT IN DROPDOWN
        Continue For Loop If    '${result}' == 'FAIL'
        
        # Server exists in dropdown. Start it and read the displayed status text.
        Click Button    xpath=//input[@value="Start Server"]
        Sleep    1s
        ${status}=    Get Text    id=status
        Save Server Status    ${server_code}    ${status}
    END

Close The Bot
    # Closes the browser after the run is complete.
    Close Browser

*** Tasks ***
Start The BotsDNA Server Project
    Open BotsDNA And Configure Downloads
    Download The Excel Input File
    Process All Servers From Python
    Send Status Email    ${MAX_EMAIL_ROWS}
    Close The Bot
