# basic_functions.psm1
Some basic functions for PowerShell that may be used in many scripts

### How to use

    Import-Module "<path_to_folder_with_module>\basic_functions.psm1" -DisableNameChecking
    
### Functions list

#### Remove-File
This function removes file if it's exist.

    Get-Help Remove-File -Full

#### Convert-DecToSysNum
This function converts decimal to any other system of numeration (up to 36).

    Get-Help Convert-DecToSysNum -Full

#### Import-ExcelAsCsv
This function imports Excel workbook to PS custom object just like Import-Csv. It could import only one selected worksheet or all of them to a single object.

    Get-Help Import-ExcelAsCsv -Full
 
#### Create-ExcelOfCSV
This function creates one Excel workbook out of one or multiple csv files. Worksheets of workbook would be named as csv files.

    Get-Help Create-ExcelOfCSV -Full

#### Get-ADCredentials
This function would show window for credentials input and would check credentials for validity.

    Get-Help Get-ADCredentials -Full

#### New-ADComputer-ADSI
This function would create computer object in specified AD OU. Also if AddUserName is specified, then all priviligies on created computer object would be granted to specified user.

    Get-Help New-ADComputer-ADSI -Full
