*** Settings ***
Documentation      Email and Excel Automation
...        Task on Email and Excel automation. 
...    An Excel file consists of a master sheet with product preference details. 
...    Needs to send an email with the attachment to multiple users and get their preference. 
...    The user has to send the attachment back after choosing YES / NO in the excel column. 
...    Once we receives the email, have to extract their feedback and needs to update the same into master sheet.


Library    RPA.Excel.Files
Library    Collections
Library    RPA.FileSystem
Library    RPA.Outlook.Application
Library    RPA.Tables
Library    RPA.RobotLogListener



*** Variables ***

${excel path}    C:/Users/priyanka.madugula/Downloads/Copy of LOU.xlsx 
${folder path}    C:/Users/priyanka.madugula/Documents/visual studio code/split sheets/
${files}    
${Account_Name}    priyanka.madugula@yash.com
${save attachment path}   C:/Users/priyanka.madugula/Documents/visual studio code/save attachment/

*** Keywords ***
Opens the Workbook
    
    TRY
        Does File Exist    ${excel path}
        ${status}=    Is File Not Empty    ${excel path}
    FINALLY
        Log    ${status}
        Open Workbook    path=${excel path}
    END   
Opens the outlook
    RPA.Outlook.Application.Open Application  

Creating multiple workbook sheets

    Open Workbook    ${excel path} 
    @{sheets}=    List Worksheets 
    FOR    ${sheet}    IN    @{sheets} 
            Set Active Worksheet    ${sheet} 
            ${table}=    Read Worksheet As Table    ${sheet}    
            Create Workbook    path=${folder path}${sheet}.xlsx    sheet_name=${sheet}
            Append Rows To Worksheet    ${table}
            Save Workbook
            Close Workbook
            Open Workbook    ${excel path}               
    END


Sends email to the users

    Open Workbook    ${excel path} 
    ${table}=    Read Worksheet As Table    Sheet2    header=True 
    ${files}=    List Files In Directory    ${folder path} 
    FOR        ${row}    IN        @{table}     
        
        Log    ${folder path}/${row}[Name].xlsx
        Log   ${row}
        Opens the outlook
        Send Message  ${row}[Email]    subject=Hi   body=Please fill the excel sheet and send feedback to the same mail   attachments=${folder path}/${row}[Name].xlsx
         
    END


Reading the user feedback attachment through mail
     
         RPA.Outlook.Application.Open Application

    Open Workbook    ${excel path}

    ${table}=    Read Worksheet As Table    Sheet2   header=True

    FOR    ${value}    IN    @{table}
            TRY
                ${waiting}    Wait For Email    ${value}[Email]    timeout=100
            EXCEPT    
                Log       Mail attachment is not found
        
            FINALLY

                Get Emails       account_name=${Account_Name}    
                ...    folder_name=Inbox
                ...    save_attachments=True
                ...    sort=True
                ...    email_filter=[Subject]='Hi'
                ...     attachment_folder=${save attachment path}
            END
            BREAK

    END
       


Getting cell value and filling the master sheet column

     ${row}=     Set Variable    2
    ${files}=    List Files In Directory    path=${save attachment path}    
     Log To Console    ${files}
    FOR    ${File}   IN    @{files}

            Open Workbook    ${File}

            @{sheets}=    List Worksheets
           
        FOR    ${sheet}    IN    @{sheets}
                
                Set Active Worksheet    ${sheet}

                ${Cell_Value}=    Get Cell Value    2    F

                Log To Console    Getting Value${Cell_Value}

                Close Workbook

                Open Workbook    ${excel path}

                Set Active Worksheet   Sheet1

                Set Cell Value    ${row}    F    ${Cell_Value} 

               # Log To Console    ${Cell_Value} 
                ${row}=    Evaluate   ${row}+1
                Log To Console    row:${row}
                Save Workbook

               

        END

       

    END

Exception
    Log    Exception Occurs Please Check the Logs

*** Tasks ***

Feedback Survey Report: Creates the Workbooks and sends to the user
    
     Opens the Workbook
    TRY
        Creating multiple workbook sheets
        Sends email to the users
    EXCEPT   
         Exception
        
    END

Reads the mail inbox to get user respone attachment and stores in the master sheet'
     Opens the outlook
    TRY

        Reading the user feedback attachment through mail
    EXCEPT   
         Exception
        
    FINALLY
        Getting cell value and filling the master sheet column
    END