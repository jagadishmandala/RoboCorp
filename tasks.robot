
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
Library    Collections

*** Variables ***
${i}=    1
${J}=    1

*** Tasks ***

Search Resourse 
    ${ConfigValues}=    Get Data From JSON File
    CheckInputFile    ${ConfigValues}[0]
    #${Se.Sheetname}=    GetSheetname
    FOR    ${i}    IN RANGE    1000
        Log    ${i}
        ${Se.Sheetname}=    GetSheetname
        IF  '${Se.Sheetname}' == '${ConfigValues}[1]'
        Failure dialog
        Log    Please Select Valid Worksheet name
        #${i}=    ${i}1
        ELSE
            #Exit For Loop If    '${Se.Sheetname}' != '${ConfigValues}[1]'
            BREAK    
        END
    END
    FOR    ${J}    IN RANGE    1000
        Log    ${J}
        ${Re.ResourceName}=    GetResourceName
        ${Lengthtable}=    Excel Operations     ${Se.Sheetname}    ${Re.ResourceName}    ${ConfigValues}[2]  
        IF  ${Lengthtable}== 0
        Failure Resource Name
        Log    Entered Resource Name is not available
        #${J}=   ${J}+    1 
        ELSE
            #Exit For Loop If    ${Lengthtable}== 0
            BREAK
        END
    END

*** Keywords ***
Failure dialog
    Add icon      Failure    
    Add heading   Invalid Sheet Selection
    Add text      Please Select the Valid Worksheet Name 
    #Show dialog    Failure   400    700    ${True}  
    Add submit buttons    buttons=OK    default=OK
    Run dialog     title=Failure    height=400    width=700 
Failure Resource Name
    Add icon      Failure    
    Add heading   Entered Invalid Resource Name
    Add text      Please Enter the Valid Resource Full Name 
    #Show dialog    Failure   400    700    ${True}  
    Add submit buttons    buttons=OK    default=OK
    Run dialog     title=Failure    height=400    width=700 
Get Data From JSON File
    ${JSONFile}    Load JSON from file    Config.json
    ${Filepath}    Get value from JSON    ${JSONFile}    $.Filepath
    ${ColumnName}    Get value from JSON    ${JSONFile}    $.ColumnName
    ${InvalidSheet}    Get value from JSON    ${JSONFile}    $.InvalidSheet
    RETURN    ${Filepath}    ${InvalidSheet}    ${ColumnName}

CheckInputFile   
    [Arguments]    ${FilePath}
    File Should Exist   ${FilePath}    msg=Excel File is not available
    TRY
    Does File Exist    ${FilePath}
    ${status}=    Is File Not Empty    ${FilePath}
    FINALLY
    Log    ${status}
    Open Workbook    path=${FilePath}
    END   
GetSheetname
    Add drop-down    Sheetname    options=Please select Worksheet,Capabilities Building,Use Cases with POC,Resource Availability,Dec Preparations    default=Please select Worksheet    
    ${s}=    Show dialog    Worksheet Name    400    700    ${True}
    ${Se}=    Wait dialog    ${s}
    Log    ${Se.Sheetname}
    RETURN    ${Se.Sheetname}
GetResourceName
    Add text input    ResourceName    label=Please Enter Resource Name
    ${s}=    Show dialog     Resource Name    400    700    ${True}
    ${Re}=    Wait dialog    ${s}
    Log    ${Re.ResourceName}    
    RETURN    ${Re.ResourceName}

Excel Operations
    [Arguments]    ${Sheetname}   ${ResourceName}    ${ColumnName} 
    Delete Columns    D
    Delete Columns    D
    Delete Columns    D
    #Save Workbook 
    ${table}=    Read Worksheet As Table    ${Sheetname}    header=True
    Filter Table By Column    ${table}    ${ColumnName}    ==    ${ResourceName} 
    ${Lengthtable}=    Get Length    ${table} 
    IF    ${Lengthtable} == 1     
    Write table to CSV    ${table}    Output.CSV
    FOR    ${element}    IN    @{table}
        Log    ${element}
        Add text    ${element}   
        ${Name}=    Set Variable    ${element}[S.No]
        Log    ${Name}
    END 
    #${Name}=    Get From Dictionary    ${table} S.No
    Log    ${Name}
    ${table}=    Show dialog     Resource-Details    400    700    ${True}
    Wait dialog    ${table} 
    Close workbook  
    ELSE
        Log    Entered Resource Name is not available
    END
    RETURN    ${Lengthtable}