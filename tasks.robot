*** Settings ***
Documentation       This is the Orders Robot created for the advanced Robot Creation Course
...                 Enters all orders from Excel file
...                 Saves the order receipt as PDF
...                 also saves the order robot as a screenshot
...                 embeds the picture in the PDF
...                 Creates a ZIP file of the PDF and the image
...
...                 TODO: use the vault and store it in the project repo
...                 TODO: Get some user input and use it
...                 TODO: add to Github
...                 TODO: Robot should be runable without local setup

Library             RPA.HTTP
Library             RPA.FileSystem
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Tables
Library             RPA.JavaAccessBridge
Library             RPA.PDF
Library             RPA.RobotLogListener
Library             RPA.Archive
Library             DateTime
Library             RPA.Robocorp.Vault


*** Variables ***
${DOWNLOAD_PATH}    ${OUTPUT DIR}${/}downloads${/}
${FILE_NAME}        orders.csv
${ORDERS_CSV}       https://robotsparebinindustries.com/orders.csv
${URL}              https://robotsparebinindustries.com/#/robot-order
${MY_OUTPUT_DIR}    ${OUTPUT_DIR}${/}output


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Create directories
    ${orders}=    Get orders
    Open the robot order website
    Close the annoying modal
    # for testing only - did not want to go through all 20 orders
    # ${counter}=    Set Variable    0

    FOR    ${row}    IN    @{orders}
        Close the annoying modal
        Fill the form    ${row}
        Preview the robot
        Submit the order
        ${pdf}=    Store the receipt as a PDF file    ${row}[Order number]
        ${screenshot}=    Take a screenshot of the robot    ${row}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
        Go to order another robot
        # for testing only - did not want to go through all 20 orders
        # ${counter}=    Evaluate    ${counter} + 1
        # IF    ${counter} == 3    BREAK
    END
    Create a ZIP file of the receipts
    [Teardown]    Log out and close the browser


*** Keywords ***
Create directories
    # Create the directory if it does not exsist
    Create Directory    ${DOWNLOAD_PATH}    exist_ok=True
    Create Directory    ${MY_OUTPUT_DIR}    exist_ok=True

Get orders
    # Connect to the online vault, must be enabled with
    # Robocorp: Connect to online secrets vault from command palette

    ${secret}=    Get Secret    orders_file_link
    TRY
        Download    url=${secret}[link]    target_file=${DOWNLOAD_PATH}    overwrite=True
    EXCEPT    PermissionError
        Log    Seems like the file was locked.... Giving up
    FINALLY
        Log out and close the browser
    END

    ${table}=    Read table from CSV    ${DOWNLOAD_PATH}${FILE_NAME}
    RETURN    ${table}

Open the robot order website
    Open Available Browser    url=${URL}

Close the annoying modal
    ${button}=    Is Element Visible
    ...    css:#root > div > div.modal > div > div > div > div > div > button.btn.btn-dark
    Log Variables
    IF    ${button} == True
        Click Button    OK
    ELSE
        Log    Button was not available
    END

Fill the form
    [Arguments]    ${row}
    Select From List By Value    id:head    ${row}[Head]
    Select Radio Button    body    ${row}[Body]
    Input Text    xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${row}[Legs]
    Input Text    id: address    ${row}[Address]
    Log    Filling the form

Preview the robot
    Click Button    id:preview

Submit the order
    Click Button    id:order

Store the receipt as a PDF file
    [Arguments]    ${order_number}
    # sometime there is an error message, let us try until it succeeds
    ${server_error}=    Set Variable    True
    WHILE    ${server_error}
        TRY
            ${sales_results_html}=    Get Element Attribute    id:receipt    outerHTML
            ${server_error}=    Set Variable    False
        EXCEPT
            Log    Server error encountered, trying again...
            Submit the order
            ${server_error}=    Set Variable    True
        END
    END
    Html To Pdf    ${sales_results_html}    ${MY_OUTPUT_DIR}${/}sales_results_${order_number}.pdf
    RETURN    ${MY_OUTPUT_DIR}${/}sales_results_${order_number}.pdf

Take a screenshot of the robot
    [Arguments]    ${order_number}
    Screenshot    id:robot-preview-image    ${MY_OUTPUT_DIR}${/}robot_order_${order_number}.png
    RETURN    ${MY_OUTPUT_DIR}${/}robot_order_${order_number}.png

Go to order another robot
    Click Button    id: order-another

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${screenshot}    ${pdf_file}
    Open Pdf    ${pdf_file}
    ${files}=    Create List
    ...    ${pdf_file}
    ...    ${screenshot}
    Add Watermark Image To Pdf    ${screenshot}    ${pdf_file}
    # Add Files To Pdf    ${files}    ${pdf_file}
    Close All Pdfs
    Remove File    ${screenshot}

Create a ZIP file of the receipts
    #${date}=    Get Current Date    result_format=YYYY-MM-DD hh:mm:ss
    ${date}=    Get Current Date    result_format=%Y-%m-%d %H-%M-%S    exclude_millis=True
    ${archive_name}=    Catenate    receipts    ${date}    .zip
    Archive Folder With Zip    ${MY_OUTPUT_DIR}    ${MY_OUTPUT_DIR}${/}${archive_name}    include=*.pdf
    ${files_to_be_removed}=    Find Files    ${MY_OUTPUT_DIR}${/}*.pdf
    FOR    ${file}    IN    @{files_to_be_removed}
        Remove File    ${file}
    END

Log out and close the browser
    Close Browser
