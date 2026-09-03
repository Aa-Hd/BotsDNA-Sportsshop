*** Settings ***
Documentation    This bot opens https://botsdna.com/sportshop/, downloads the School Email
...    Database,scrapes the latest product list (name/code/unit price), 
...    reads the list of schools from the Schools page, then for each school:
...    searches for that school on the database to reveal its discount % and get school's mail,
...    builds a "Price Quotation For Various New document (same layout as SportsTemplet.pdf) with offer prices computed
...    from the discount, and emails it to that school's address.
...    Sends a summary email at the end to admin, and a failure email if the run breaks.

Library      Autosphere.Browser.Selenium
Library      Autosphere.Desktop.OperatingSystem
Library      OperatingSystem
Library      Autosphere.HTTP
# Library      Autosphere.Word.Application
Library      config_reader.py
Library      String
Library      Collections
Library      DateTime
Library      Autosphere.Email.ImapSmtp   smtp_server=smtp.gmail.com  smtp_port=587
Library      Autosphere.Excel.Files
Library      Autosphere.RobotLogListener
Library      dox_handler.py
Library      Autosphere.Tasks
Library      Autosphere.FileSystem
Library      popup.py
# Library    Autosphere.Excel.Application
# Suite Teardown    Run Keyword If All Tests Passed    Send Summary Email

*** Keywords ***
Read Config File
    ${config}=    Get Config Variables    filepath=D:\\Sport_Shop_Bot\\config.ini
    # Log    ${config}
    ${current_date}=  Get Current Date   result_format=%d-%m-%Y
    Set Global Variable    ${config}
    Set Global Variable    ${current_date}
    IF  ${config}== {}
        Log    ERROR: Config file not found or failed to read config file.\n Please check the config file and try again.
        Jump To Task   Process End
    END

Open Sportshop Website
    ${status}=    Kill Process   chrome.exe
    ${browser}=  Run Keyword And Return Status    Autosphere.Browser.Selenium.Open Browser    ${config}[sportshop_url]
    Run Keyword And Ignore Error   Autosphere.Browser.Selenium.Maximize Browser Window
    IF  ${browser}==False
        ${Email_Body}=  Set Variable  ERROR:Cannot Open Browser ,kindly check relevant dependencies
        Send Email Notification    ${Email_Body}
        Jump To Task    Process End
    ELSE
    ${download_status}=  Run Keyword And Return Status  Download Reference Files
    ${products_status}=  Run Keyword And Return Status  Get Latest Products
    ${schools_status}=  Run Keyword And Return Status  Get Schools Strength
    
        IF  (${download_status} and ${products_status} and ${schools_status})!= True
            Jump To Task    Process End                            
        END
    END
Download Reference Files
    # Run Keyword If File Exists   ${config}[template_path]  OperatingSystem.Remove File  ${config}[template_path]
    # Run Keyword If File Exists   ${config}[email_database_path]  OperatingSystem.Remove File  ${config}[email_database_path]
    OperatingSystem.Remove Files    ${config}[template_path]  ${config}[email_database_path]
    ${template_file}=  Run Keyword And Return Status  Download    ${config}[template_url]   target_file=${config}[template_path]
    ${db_status}=  Run Keyword And Return Status
    ...    Download    ${config}[email_database_download_link]    target_file=${config}[email_database_path]
    ${db_created}=  Run Keyword And Return Status
    ...    OperatingSystem.Wait Until Created    ${config}[email_database_path]    timeout=60s
    IF  '${template_file}'=='False'
        ${Email_Body}=  Set Variable   ERROR: Failed to Download Quotation Template File: "SportsTemplet.docx"
    ELSE IF  (${db_status} or ${db_created})==False
        ${Email_Body}=  Set Variable    ERROR: Failed to download EmailsDatabase.xlsx from the sportshop site.\n Please check the URL / network and try again.
        Send Email Notification    ${Email_Body}
    END
Get Product Details
    [Arguments]          ${i}
    ${table_xpath}=    Set Variable
    ...    (//span[contains(@style,'color:red')]/ancestor::table[starts-with(@id,'tbl')])[${i}]

    ${product_name}=    Autosphere.Browser.Selenium.Get Text
    ...    ${table_xpath}//tr[td[1][contains(normalize-space(),'Product Name')]]/td[2]

    ${product_code}=    Autosphere.Browser.Selenium.Get Text
    ...    ${table_xpath}//tr[td[1][contains(normalize-space(),'Product Code')]]/td[2]

    ${product_price}=    Autosphere.Browser.Selenium.Get Text
    ...    ${table_xpath}//tr[td[1][contains(normalize-space(),'Price')]]/td[2]

    ${product_arrived_from}=    Autosphere.Browser.Selenium.Get Text
    ...    ${table_xpath}//tr[td[1][contains(normalize-space(),'Stock Arrived From')]]/td[2]
    
    ${product_img_path}=  Set Variable   ${config}[screenshot_path]${i}.jpg
    ${product_img}=  Screenshot   ${table_xpath}//tr/td/img   filename=${product_img_path}
    #(//span[contains(@style,'color:red')]/ancestor::table[starts-with(@id,'tbl')])[1]//tr/td/img
    RETURN   ${product_name}    ${product_code}     ${product_price}    ${product_arrived_from}   ${product_img_path}
Get Latest Products 
        Wait Until Page Contains Element    //span[contains(@style,'color:red')]/ancestor::table[starts-with(@id,'tbl')]     timeout=120
        
        ${latest_products_count}=  Autosphere.Browser.Selenium.Get Element Count
    ...    //span[contains(@style,'color:red')]/ancestor::table[starts-with(@id,'tbl')]
        ${latest_products}=    Create List
    IF   '${latest_products_count}' != '0'
        FOR    ${i}    IN RANGE    1    ${latest_products_count}+1
            ${product_name}    ${product_code}     ${product_price}    ${product_arrived_from}    ${product_img_path}=  Get Product Details    ${i}       
            ${product_price}=  Split String  ${product_price}   /
            ${product_price}=  Strip String  ${product_price}[0]

            &{product}=    Create Dictionary
            ...    Product No.=${i}
            ...    Product Name=${product_name}
            ...    Product Code=${product_code}
            ...    Price=${product_price}
            ...    Stock Arrived From=${product_arrived_from}
            ...    Product Image=${product_img_path}

            Append To List    ${latest_products}    ${product}
            Log    ${latest_products}
            Set Suite Variable    ${latest_products}
            Set Suite Variable    ${latest_products_count}
        END
        
    ELSE
        Log  No Latest Products Found
        ${Email_Body}=  Set Variable   No Latest Products Found For Today.So,Ending Process at this step
        Send Email Notification    ${Email_Body}
        Skip
    END
    
Get Schools Strength
    # Autosphere.Browser.Selenium.Open Browser    ${config}[schools_page_url]
    # Autosphere.Browser.Selenium.Maximize Browser Window
    #Autosphere.Browser.Selenium.Go To    ${config}[schools_page_url]
    ${schools_page}=  Run Keyword And Return Status   Autosphere.Browser.Selenium.Click Element    //table//td//a[@href="schools.html"]
    IF  ${schools_page}
        @{all_schools}=    Create List      
        ${row_count}=   Autosphere.Browser.Selenium.Get Element Count     ${config}[schools_table_row]

        FOR    ${i}    IN RANGE    1    ${row_count}
            ${school_code}=    Autosphere.Browser.Selenium.Get Text     (//table[@id='courts']//tr[td])[${i}]/td[1]

            ${school_name}=    Autosphere.Browser.Selenium.Get Text     (//table[@id='courts']//tr[td])[${i}]/td[2]

            ${student_strength}=    Autosphere.Browser.Selenium.Get Text     (//table[@id='courts']//tr[td])[${i}]/td[3]

            &{school}=    Create Dictionary
            ...    School Code=${school_code}
            ...    School Name=${school_name}
            ...    Student Strength=${student_strength}

            Append To List    ${all_schools}    ${school}
        
        END
        Log    ${all_schools} 
        Set Suite Variable    ${all_schools}
    ELSE
        ${Email_Body}=  Set Variable  ERROR: Schools page cannot be accessed.Cannot move further ,so ending process here.
        Send Email Notification    ${Email_Body}
    END
    Close Browser
Read Discounts And Schools Mails List
    ${workbook}=  Run Keyword And Return Status  Open Workbook    ${config}[email_database_path]
    ${discount_offers}=  Create Dictionary
    ${schools_mails}=  Create Dictionary
    IF  ${workbook}  
         @{list_of_worksheets}=  List Worksheets
        FOR   ${sheet}   IN   @{list_of_worksheets} 
        
            IF  '${sheet}'=='Offers'
            Set Active Worksheet    Offers
            ${offers}=  Read Worksheet    header=${True}
            ${discount_offers}=  Set Variable   ${offers}

            ELSE IF  '${sheet}'=='SchoolsData'
            Set Active Worksheet    SchoolsData
            ${mails}=  Read Worksheet     header=${True}
            ${schools_mails}=  Set Variable    ${mails}
            END
        
        END
    Close Workbook
    ELSE
        ${Email_Body}=  Set Variable  ERROR:Workbook Cannot be Opened.
        Send Email Notification    ${Email_Body}
        Jump To Task    Process End
    END
    Set Suite Variable      ${discount_offers}  
    Set Suite Variable      ${schools_mails}
    Log Many   ${discount_offers}  ${schools_mails}
    Close Workbook
    RETURN    ${discount_offers}   ${schools_mails}
Get Email of School 
    [Arguments]    ${school}
    ${school_code}=  Strip String  ${school}[School Code]
    ${school_mail}=  Set Variable  ${EMPTY}
    FOR    ${record}    IN    @{schools_mails}
        ${school_code_in_mails}=  Strip String  ${record}[School Code]
        IF    '${school_code}' == '${school_code_in_mails}'
            ${pt_mail}=  Set Variable   ${record}[PT Email]
            ${admin_mail}=  Set Variable   ${record}[Administrator Email]
            ${principal_mail}=  Set Variable    ${record}[Principal Email]
            ${director_mail}=  Set Variable   ${record}[Director Email]
            IF   ('${pt_mail}' != '${EMPTY}') and ('${pt_mail}' != 'None')
                ${pt_mail}=  Strip String  ${record}[PT Email]
                ${school_mail}=  Set Variable  ${pt_mail}
            
            ELSE IF   ('${admin_mail}' != '${EMPTY}') and ('${admin_mail}' != 'None')
                ${admin_mail}=  Strip String  ${record}[Administrator Email]
                ${school_mail}=  Set Variable  ${admin_mail}
            
            ELSE IF   ('${principal_mail}' != '${EMPTY}') and ('${principal_mail}' != 'None')
                ${principal_mail}=  Strip String    ${record}[Principal Email]
                ${school_mail}=  Set Variable  ${principal_mail}

            ELSE IF   ('${director_mail}' != '${EMPTY}') and ('${director_mail}' != 'None')
                ${director_mail}=  Strip String   ${record}[Director Email]
                ${school_mail}=  Set Variable  ${director_mail}

            END
            Log    ${school_mail}
        END
    END
    Log   ${school_mail}
    RETURN    ${school_mail}
Extract Discount Prices of Products
    [Arguments]        ${offer}    #makes discounted prices of all products ,and updates to 
    ${discounted_products}=   Copy List   ${latest_products}
    FOR   ${product}  IN  @{latest_products}
        ${price}=  Set Variable  ${product}[Price]    #changed
        ${discount_price}=  Evaluate  ${price}*${offer}
        ${price_after_discount}=  Evaluate  ${price}-${discount_price}
        Set To Dictionary    ${product}    Discounted Price=${price_after_discount}
        Log    ${product}
        Log    ${discounted_products}
    END   
    RETURN    ${discounted_products}
Get Discount For School 
    [Arguments]      ${school}
    # FOR  ${school}   IN   @{all_schools}
    ${school_name}=  Set Variable  ${school}[School Name]
    ${school_code}=  Set Variable  ${school}[School Code]
    ${school_strength}=  Convert To Integer  ${school}[Student Strength]

    ${discounted_products}=  Set Variable   ${EMPTY}
    ${school_discount}=  Set Variable   ${EMPTY}

    FOR   ${criteria}  IN   @{discount_offers}
        ${size_from}=    Convert To Integer    ${criteria}[Size From]
        ${size_to}=      Convert To Integer    ${criteria}[Size To]
        ${discount}=     Set Variable   ${criteria}[Offer %]      #changed
        ${offer}=   Evaluate  ${discount}*0.01
        
        IF  (${school_strength} >= ${size_from}) and (${school_strength} <= ${size_to})
            ${discounted_products}=  Extract Discount Prices of Products    ${offer} 
            ${school_discount}=  Set Variable  ${discount}
        END
    END

    IF  '${school_discount}'=='${EMPTY}'
        
        ${total_offers}=  Get Length  ${discount_offers}
        ${last_offer_index}=  Evaluate  ${total_offers}-1
        ${last_offer}=  Set Variable  ${discount_offers}[${last_offer_index}]    #index of last item in discounts,which is a max discount
        ${highest_discount}=  Set Variable  ${last_offer}[Offer %]     #changed
        ${highest_offer}=   Evaluate  ${highest_discount}*0.01
        ${discounted_products}=  Extract Discount Prices of Products    ${highest_offer}
        ${school_discount}=  Set Variable  ${highest_discount}
    END
    Log Many       ${discounted_products}    ${school_discount}
    RETURN     ${discounted_products}    ${school_discount}
# Create Quotation PDF
#     [Arguments]     ${school}    ${discounted_products_for_school}   ${school_discount}
#     Open Application
#     Open File    ${config}[template_path]
#     Write Text    \nYour School ${school}[School Name] Has ${school_discount}% Discount on All Products  end_of_text=True
#     FOR  ${product}  IN  @{discounted_products_for_school}    
        
#         Write Text    \nn#${product}[Product No.]                         end_of_text=True
#         Write Text    \n${config}[screenshot_path]${product}[Product No.].jpg     end_of_text=True
#         Write Text    \nPRODUCT NAME: ${product}[Product Name]            end_of_text=True
#         Write Text    \nPRODUCT CODE: ${product}[Product Code]            end_of_text=True
#         Write Text    \nREGULAR PRICE: ${product}[Price]/_                end_of_text=True
#         Write Text    \nDISCOUNTED PRICE: ${product}[Discounted Price]/_  end_of_text=True
#         Write Text    \nSTOCK ARRIVED FROM: ${product}[Stock Arrived From]  end_of_text=True

#     END
#     Create Directory    ${config}[quotation_path]
#     ${quotation_file_path}=  Set Variable  ${config}[quotation_path]\\${school}[School Name]_Sports_Quotation.pdf
#     Export To Pdf    ${quotation_file_path} 
#     Quit Application
#     Set Suite Variable   ${quotation_file_path}
#     RETURN     ${quotation_file_path}
Create Quotation PDF
    [Arguments]     ${school}    ${discounted_products_for_school}    ${school_discount}

    # OperatingSystem.Create Directory    ${config}[quotation_path]   #simply pass if directory already exists
    ${docx_path}=    Set Variable    ${config}[quotation_path]\\${school}[School Name]_Sports_Quotation.docx
    ${quotation_file_path}=    Set Variable    ${config}[quotation_path]\\${school}[School Name]_Sports_Quotation.pdf

    ${open_template_docx}=  Copy Docx    ${config}[template_path]      ${docx_path}
    ${top_line}=  Write Text To Docx    ${docx_path}    ${docx_path}
        ...    Your School ${school}[School Name] Has ${school_discount}% Discount on All Products${\n}

    FOR    ${product}    IN    @{discounted_products_for_school}
        Write Text To Docx    ${docx_path}    ${docx_path}    n#${product}[Product No.]

        ${image_path}=    Set Variable    ${config}[screenshot_path]${product}[Product No.].jpg
        Add Picture To Docx    ${docx_path}    ${docx_path}    ${image_path}

        Write Text To Docx    ${docx_path}    ${docx_path}    PRODUCT NAME: ${product}[Product Name]
        Write Text To Docx    ${docx_path}    ${docx_path}    PRODUCT CODE: ${product}[Product Code]
        Write Text To Docx    ${docx_path}    ${docx_path}    REGULAR PRICE: ${product}[Price]/_
        Write Text To Docx    ${docx_path}    ${docx_path}    DISCOUNTED PRICE: ${product}[Discounted Price]/_${\n}
        # Write Text To Docx    ${docx_path}    ${docx_path}    STOCK ARRIVED FROM: ${product}[Stock Arrived From]
        
    END
    ${pdf_status}=  Run Keyword And Return Status  Convert Docx To Pdf    ${docx_path}    ${quotation_file_path}
    # Set Suite Variable    ${quotation_file_path}
    IF  '${pdf_status}'=='False'
   
        ${Email_body}=  Set Variable  ERROR:Quotation Pdf File Not Created.${KEYWORD_MESSAGE}
        Send Email Notification    ${Email_Body}
    END
    RETURN    ${quotation_file_path}
Process
    OperatingSystem.Create Directory    ${config}[quotation_path]   #simply pass if directory already exists
    FOR  ${school}   IN   @{all_schools}
        Show Popup    Quotation ${school}[School Name] execution is running!.
        ${discounted_products_for_school}  ${school_discount}=  Get Discount For School    ${school}
        ${school_mail}=  Get Email of School    ${school}
        ${quotation_file_path}=  Create Quotation PDF   ${school}   ${discounted_products_for_school}   ${school_discount}
        ${status}=  Run Keyword And Return Status  Send Quotation Email    ${school_mail}    ${school}[School Name]    ${school_discount}    ${quotation_file_path}
        IF  '${status}'=='False'
            ${Email_Body}=  Set Variable    Quotation Sending to ${school}[School Name] Failed.
            Show Popup    Quotation Sending to ${school}[School Name] failed.
            Send Email Notification    ${Email_Body}
        ELSE
            Log    Quotaion Sent to ${school}[School Name] Successfully.
            Show Popup   Quotation Sent to ${school}[School Name] Successfully.
        END
    END    
Send Email Notification
    [Arguments]     ${Email_Body}
    Authorize  account=${config}[sender_gmail]  password=${config}[gmail_password]
    Send Message
        ...    sender=${config}[sender_gmail]
        ...    recipients=${config}[reciever_gmail]
        ...    subject=Sports Shop Bot Update
        ...    body=Dear Concern,${\n}${Email_Body}${\n}Regards,${\n} RPA BOT
Send Quotation Email
    [Arguments]    ${recipient}    ${school_name}    ${discount_percent}    ${attachment}
    Authorize  account=${config}[sender_gmail]  password=${config}[gmail_password]
    ${msg}=    Catenate    SEPARATOR=\n
    ...    It is my great pleasure to offer the best price in the latest stock. 
    ...    Please take a moment to review the sports products and get back with order.
    ...    If you have any further questions please donot hesitate to contact me.
    ...    Sincerely,
    ...    Sports Shop Team
   
    Send Message
        ...    sender=${config}[sender_gmail]
        ...    recipients=${recipient}
        ...    subject=${school_name} ${config}[quotation_mail_subject]  
        ...    body=Dear ${school_name},${\n}${msg}
        ...    attachments=${attachment}
    
Failure Mail
    ${Email_Body}=  Set Variable    ERROR: Sports Shop quotation bot execution failed. ${\n} Showing message "${TEST_MESSAGE}".${\n} Please check the logs for more details.
    Send Email Notification    ${Email_Body}
Send Summary Email
    ${Email_Body}=    Set Variable    Sports Shop quotation run complete for ${current_date}.${\n} Quotations sent successfully and Backup is done.
    Send Email Notification    ${Email_Body}
*** Tasks ***
Send Sports Shop Quotations
    Read Config File 
    Open Sportshop Website
    # Download Reference Files
    # Get Latest Products
    # Get Schools Strength
    Read Discounts And Schools Mails List
    Process
    Jump To Task    Move Files To Backup
    [Teardown]   Run Keywords
    ...     Run Keyword If Test Failed    Failure Mail    
    ...     AND
    ...     Run Keyword And Ignore Error  Autosphere.Browser.Selenium.Close Browser
Move Files To Backup
    # OperatingSystem.Empty Directory     ${config}[backup_folder]\\${current_date}
    OperatingSystem.Remove Directory    ${config}[backup_folder]\\${current_date}   recursive=True
    OperatingSystem.Create Directory    ${config}[backup_folder]\\${current_date}
    OperatingSystem.Move File    ${config}[email_database_path]   ${config}[backup_folder]\\${current_date}
    OperatingSystem.Move File    ${config}[template_path]        ${config}[backup_folder]\\${current_date}
    OperatingSystem.Move Directory   ${config}[quotation_path]   ${config}[backup_folder]\\${current_date}
    
    [Teardown]    Run Keywords
    ...    Run Keyword If Test Failed    Failure Mail
    ...    AND 
    ...    Run Keyword If Test Passed     Jump To Task    Process Completion Mail
              
Process Completion Mail
    Send Summary Email
    Jump To Task    Process End
Process End
    Log  Process Ended
      