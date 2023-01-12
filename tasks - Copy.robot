
*** Settings ***
Documentation      Find Resourece Status
...        Task on UI & Excel automation. 
...    An Excel file consists of a mMultiple sheets and with dialog box user will select the Sheet name.
...    User need to enter the Resourece name they want to find the resourece status  
...    Once we receives the Data, bot will filter data and save required data in Output.csv file.

Library    RPA.Excel.Files
Library    RPA.Dialogs
Library    RPA.Tables
Library    OperatingSystem
Library    RPA.FileSystem
Library    RPA.JSON

*** Variables ***

#${filename}    Robocorp.xlsx
#${ResourceName1}    ResourceName
#${InvalidSheet}    Please select Worksheet

*** Tasks ***

Search Resourse 
    ${ConfigValues}=    Get Data From JSON File
    ${Inputvalues}=    User Input    ${ConfigValues}[0]    ${ConfigValues}[1]
    Excel Operations     ${Inputvalues}[0]    ${Inputvalues}[1]    ${ConfigValues}[2]

*** Keywords ***

Get Data From JSON File
    ${JSONFile}    Load JSON from file    Config.json
    ${Filepath}    Get value from JSON    ${JSONFile}    $.Filepath
    ${ColumnName}    Get value from JSON    ${JSONFile}    $.ColumnName
    ${InvalidSheet}    Get value from JSON    ${JSONFile}    $.InvalidSheet
    RETURN    ${Filepath}    ${InvalidSheet}    ${ColumnName}

UserInput   
    [Arguments]    ${FilePath}   ${InvalidSheet}
    File Should Exist   ${FilePath}    msg=Excel File is not available
    TRY
    Does File Exist    ${FilePath}
    ${status}=    Is File Not Empty    ${FilePath}
    FINALLY
    Log    ${status}
    Open Workbook    path=${FilePath}
    END   
    Add drop-down    Sheetname    options=Please select Worksheet,Capabilities Building,Use Cases with POC,Resource Availability,Dec Preparations    default=Please select Worksheet    
    ${s}=    Show dialog    Search Resource    400    700    ${True}
    ${Se}=    Wait dialog    ${s}
    Add text input    ResourceName    label=Please Enter Resource Name
    ${s}=    Show dialog     Search Resource    400    700    ${True}
    ${Re}=    Wait dialog    ${s}
    Log    ${Re.ResourceName}    
    Log    ${Se.Sheetname}
    IF    '${Se.Sheetname}' == '${InvalidSheet}'
        Log    Please Select Valid Worksheet name
        
    END
    RETURN    ${Se.Sheetname}    ${Re.ResourceName}

Excel Operations
    [Arguments]    ${Sheetname}   ${ResourceName}    ${ColumnName} 
    Delete Columns    D
    Delete Columns    D
    Delete Columns    D
    #Save Workbook 
    ${table}=    Read Worksheet As Table    ${Sheetname}    header=True
    #${rows}=    Read Worksheet    Capabilities Building    True      
    #${rows}=    Filter Table With Keyword    ${table}    Resource Name
    Filter Table By Column    ${table}    ${ColumnName}    ==    ${ResourceName} 
    ${Lengthtable}=    Get Length    ${table} 
    IF    ${Lengthtable} == 1     
    Write table to CSV    ${table}    Output.CSV
    FOR    ${element}    IN    @{table}
        Log    ${element}
        Add text    ${element}
        ${ReadDialog}=    Show dialog       
        Wait dialog    ${ReadDialog}
    END 
    Close workbook  
    ELSE
        Log    Entered Resource Name is not available
    END