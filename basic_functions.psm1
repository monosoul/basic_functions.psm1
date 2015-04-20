Function Remove-File {
  <#
    .Synopsis
       Removes file if it's exist
    .DESCRIPTION
       The function removes file if it's exist.
    .EXAMPLE
       Remove-File "C:\Temp\temp.txt"
    .PARAMETER fileName
      Path to file
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Removes file if it's exist
  #>
  param(
    [parameter(Mandatory = $true)][string]$fileName
  )
  if(Test-Path -path $fileName) { Remove-Item -path $fileName -Force -Confirm:$false }
}

Function Convert-DecToSysNum {
  <#
    .Synopsis
       Converts decimal to any other system of numeration.
    .DESCRIPTION
       The function converts decimal to any other system of numeration.
    .EXAMPLE
       Convert-DecToSysNum -number 795 -numsystem 2

       Result:
       1100011011
    .EXAMPLE
       Convert-DecToSysNum -number 795 -numsystem 16

       Result:
       31B
    .EXAMPLE
       Convert-DecToSysNum -number 795 -numsystem 26 -onlyalphabet

       Result:
       BEP
    .PARAMETER number
      Number to convert
    .PARAMETER numsystem
      System of numeration you want to use
    .PARAMETER onlyalphabet
      Use this switch if you want to use only alphabetic system of numeration
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Converts decimal to any other system of numeration.
  #>
  param(
    [parameter(Mandatory = $true)][int64]$number,
    [parameter(Mandatory = $true)][int]$numsystem,
    [switch]$onlyalphabet
  )
  if ($numsystem -gt 36) {
    Write-Error "System of numeration could only consist of 36 symbols!"
    return $null
  }
  [string]$result = ""
  [string]$alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Do {
    $full = [math]::Truncate($number / $numsystem)
    $left = $number % $numsystem
    if ($onlyalphabet.IsPresent) {
      if ($numsystem -gt 26) {
        Write-Error "Alphabetic system of numeration could only consist of 26 symbols!"
        return $null
      }
      [string]$result = "$($alphabet[$left])$result"
    } else {
      if ($left -gt 9) {
        [string]$result = "$($alphabet[$left - 10])$result"
      } else {
        [string]$result = "$left$result"
      }
    }
    $number = $full
  } while (($left -gt ($numsystem - 1)) -or ($full -gt 0))
  return $result
}

Function Import-ExcelAsCsv {
  <#
    .Synopsis
       Import Excel workbook to PS custom object just like csv.
    .DESCRIPTION
       The function imports Excel workbook to PS custom object just like Import-Csv.
       It could import only one selected worksheet or all of them to a single object.
    .EXAMPLE
       $tmpObj = Import-ExcelAsCsv -FileName "C:\Temp\workbook.xlsx" -noheaders

       Result:
       Contents                                                                       Name                                                                          
       --------                                                                       ----                                                                          
       {@{A=arg; B=aer}, @{A=are g; B=g areg}}                                        List1                                                                         
       {@{A=as d32 r2; B=123 rd ; C=ed32; D=ed2}, @{A=aw rf; B=e ed; C=asfas; D=Edd}} List2                                                                         
       {@{A=asdas; B=asdasdasdasdasd; C=asdasd}, @{A=a1312; B=12352341; C=c23cr2r}}   List3
    .EXAMPLE
       $tmpObj = Import-ExcelAsCsv -FileName "C:\Temp\workbook.xlsx" -WorksheetName "List1"

       Result:
       arg                                                                            aer                                                                           
       ---                                                                            ---                                                                           
       are g                                                                          g areg
    .PARAMETER fileName
      Path to file.
    .PARAMETER WorksheetName
      Name of worksheet to import.
    .PARAMETER noheaders
      Use this switch if you don't want to use headers from worksheets.
      Alphabetic system of numeration would be used then instead.
    .NOTES
       You would need MS Office installed for this function to work.

       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Import Excel workbook to PS custom object just like csv.
  #>
  param (
    [parameter(Mandatory = $true)][string]$FileName,
    [string]$WorksheetName,
    [switch]$noheaders
  )
  if(Test-Path -path $FileName) {
   
  
   $excelObject = New-Object -ComObject Excel.Application  
   $excelObject.Visible = $false 
   $workbookObject = $excelObject.Workbooks.Open($filename)
   While ($excelObject.Ready -ne $true) {
     Start-Sleep 1
   }
   if ($WorksheetName) {
     $csvFile = ($env:temp + "\" + ((Get-Item -path $FileName).name).Replace(((Get-Item -path $FileName).extension),".csv"))
     Remove-File "$csvFile"
     $worksheet = $workbookObject.Sheets.item($WorksheetName)
     $worksheet.SaveAs($csvFile,6)
     if ($noheaders.IsPresent) {
       $header = @()
       1..$worksheet.UsedRange.Columns.Count | ForEach-Object {
         $header += Convert-DecToSysNum -number $($_ - 1) -numsystem 26 -onlyalphabet
       }
     }
     $workbookObject.Saved = $true
     $workbookObject.Close()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) | Out-Null
     $excelObject.Quit()
     if ($noheaders.IsPresent) {
       $spreadsheetList = Import-Csv -path "$csvFile" -Encoding Default -Delimiter ';' -Header $header
     } else {
       $spreadsheetList = Import-Csv -path "$csvFile" -Encoding Default -Delimiter ';'
     }
     Remove-File "$csvFile"
   } else {
     $spreadsheetList = @()
     $indexes = @()
     ForEach ($worksheet in $workbookObject.Sheets) {
       $sheetname = $worksheet.Name
       if ($noheaders.IsPresent) {
         $header = @()
         1..$worksheet.UsedRange.Columns.Count | ForEach-Object {
           $header += Convert-DecToSysNum -number $($_ - 1) -numsystem 26 -onlyalphabet
         }
       }
       $csvFile = ($env:temp + "\" + ((Get-Item -path $FileName).name).Replace(((Get-Item -path $FileName).extension),".csv_$($worksheet.Index)"))
       $indexes += $worksheet.Index
       Remove-File "$csvFile"
       $worksheet.SaveAs($csvFile,6)
       $spreadsheetList += New-Object PSObject -Property @{
         Name = $sheetname
         Contents = &{if ($noheaders.IsPresent) {
                        Import-Csv -path "$csvFile" -Encoding Default -Delimiter ';' -Header $header
                      } else {
                        Import-Csv -path "$csvFile" -Encoding Default -Delimiter ';'
                      }
                     }
       }
     }
     $workbookObject.Saved = $true
     $workbookObject.Close()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) | Out-Null
     $excelObject.Quit()
     $indexes | ForEach-Object {
       $csvFile = ($env:temp + "\" + ((Get-Item -path $FileName).name).Replace(((Get-Item -path $FileName).extension),".csv_$($_)"))
       Remove-File "$csvFile"
     }
   }

   [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) | Out-Null
   [System.GC]::Collect()
   [System.GC]::WaitForPendingFinalizers()

   return $spreadsheetList
  
  } else {
    Write-Error "File doesn't exist!" -TargetObject $FileName -Category ObjectNotFound
    return $null
  }
}

Function Create-ExcelOfCSV {
  <#
    .Synopsis
       Creates one Excel workbook out of one or multiple csv files.
    .DESCRIPTION
       The function creates one Excel workbook out of one or multiple csv files. Worksheets of workbook would be named as csv files.
    .EXAMPLE
       $csvlist = @()
       $csvlist += "c:\Temp\csv1.csv"
       $csvlist += "c:\Temp\csv2.csv"

       Create-ExcelOfCSV -fileslist $csvlist -resultname "c:\Temp\result.xlsx"
    .PARAMETER fileslist
      List of csv files from which xlsx going to be created.
    .PARAMETER resultname
      Path to resulting file.
    .NOTES
       You would need MS Office installed for this function to work.

       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Creates one Excel workbook out of one or multiple csv files.
  #>
  param(
    [parameter(Mandatory = $true)][PSObject]$fileslist,
    [parameter(Mandatory = $true)][string]$resultname
  )

  $excelObject = New-Object -ComObject Excel.Application
  $excelObject.Visible = $false
  $excelObject.Application.DisplayAlerts = $false

  if(Test-Path -path "$resultname") {
    $workbookObject = $excelObject.Workbooks.Open("$resultname")

    ForEach ($file in $fileslist) {
      $filename = $file -replace ".*\\",""
      $worksheetObject = $null
      # выбираем лист по имени
      try {
        $worksheetObject = $workbookObject.Worksheets.Item($($filename -replace '.csv',''))
      }
      catch {
      }
      # если лист с таким именем есть
      if ($worksheetObject) {
        # то дописываем информацию в него
        
        $TxtConnector = ("TEXT;$file")
        $Connector = $worksheetObject.QueryTables.Add($TxtConnector,$worksheetObject.Range("A$($worksheetObject.UsedRange.Rows.Count + 1)"))
        $query = $worksheetObject.QueryTables.Item($Connector.Name)
        
        ## указываем разделитель
        $query.TextFileOtherDelimiter = $excelObject.Application.International(5)

        ## указываем формат
        $query.TextFileParseType = 1
        $query.TextFileColumnDataTypes = ,1 * $worksheetObject.Cells.Columns.Count
        $query.AdjustColumnWidth = 1
        $query.Refresh() | Out-Null
        $query.Delete() | Out-Null
        $worksheetObject.Rows.Item($($worksheetObject.UsedRange.Rows.Count + 1)).Delete() | Out-Null
      } else {
        # если нету, то создаём новый лист
        $worksheetObject = $workbookObject.Worksheets.Add()

        $TxtConnector = ("TEXT;$file")
        $Connector = $worksheetObject.QueryTables.Add($TxtConnector,$worksheetObject.Range("A1"))
        $query = $worksheetObject.QueryTables.Item($Connector.Name)

        ## указываем разделитель
        $query.TextFileOtherDelimiter = $excelObject.Application.International(5)

        ## указываем формат
        $query.TextFileParseType = 1
        $query.TextFileColumnDataTypes = ,1 * $worksheetObject.Cells.Columns.Count
        $query.AdjustColumnWidth = 1
        $query.Refresh() | Out-Null
        $query.Delete() | Out-Null

        $worksheetObject.Name = $filename -replace "`.csv",""
        $worksheetObject.UsedRange.Rows.Item(1).AutoFilter() | Out-Null
      }
    }

  } else {
    $workbookObject = $excelObject.Workbooks.Add(1)

    ForEach ($file in $fileslist) {
      $filename = $file -replace ".*\\",""
      $worksheetObject = $workbookObject.Worksheets.Add()

      $TxtConnector = ("TEXT;$file")
      $Connector = $worksheetObject.QueryTables.Add($TxtConnector,$worksheetObject.Range("A1"))
      $query = $worksheetObject.QueryTables.Item($Connector.Name)

      ## указываем разделитель
      $query.TextFileOtherDelimiter = $excelObject.Application.International(5)

      ## указываем формат
      $query.TextFileParseType = 1
      $query.TextFileColumnDataTypes = ,1 * $worksheetObject.Cells.Columns.Count
      $query.AdjustColumnWidth = 1
      $query.Refresh() | Out-Null
      $query.Delete() | Out-Null

      $worksheetObject.Name = $filename -replace "`.csv",""
      $worksheetObject.UsedRange.Rows.Item(1).AutoFilter() | Out-Null
    }
    # удаляем лишний лист
    $workbookObject.Worksheets.Item((@($fileslist).Count + 1)).Delete() | Out-Null
  }
  if ($query) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) -ge 0) {}
  }
  $query = $null
  if ($Connector) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) -ge 0) {}
  }
  $Connector = $null

  try {
    $workbookObject.Names.Item("_FilterDatabase").Delete() | Out-Null
  }
  catch {
  }
  $Items_FD = $workbookObject.Names | Where-Object { $_.Name -match "_FilterDatabase" }
  $Items_FD | ForEach-Object {
    try{
      $_.Delete() | Out-Null
      while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($_) -ge 0) {}
    }
    catch{
    }
  }
  if ($Items_FD) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Items_FD) -ge 0) {}
  }
  $Items_FD = $null
  if ($worksheetObject) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetObject) -ge 0) {}
  }
  $worksheetObject = $null
  $workbookObject.SaveAs("$resultname")
  $workbookObject.Saved = $true
  $workbookObject.Close()
  if ($workbookObject) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) -ge 0) {}
  }
  $workbookObject = $null
  $excelObject.Quit()
  if ($excelObject) {
    while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) -ge 0) {}
  }
  $excelObject = $null
  Remove-Variable -Force * -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

}

Function Get-ADCredentials {
  <#
    .Synopsis
       Gets AD credentials and checks it for validity.
    .DESCRIPTION
       The function would show window for credentials input and would check credentials for validity.
    .EXAMPLE
       $objCreds = Get-ADCredentials -message "Please enter user name with administrator priviliges"
    .EXAMPLE
       $objCreds = Get-ADCredentials -message "Please enter user name with administrator priviliges" -username "Administrator"
    .PARAMETER message
      Message that would be shown in credentials input form.
    .PARAMETER username
      Pre-defined user name.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Gets AD credentials and checks it for validity.
    .OUTPUTS
       [PSCredential]
  #>
  param(
    [parameter(Mandatory = $true)][string]$message,
    [string]$username=$null
  )

  $adcreds = $null

  Write-Host $message -f Red

  while ($domain.name -eq $null) {
    $ADCredentials = Get-Credential -Message $message -UserName $username
    $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
    $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $ADCredentials.username, $ADCredentials.GetNetworkCredential().password)
    if ($domain.name -eq $null) {
      Write-Host "Введены неверные имя пользователя и пароль!!!" -f red
      $rep = Read-Host "Повторить ввод? (Y/N)"
      if (($rep -eq "N") -or ($rep -eq "n")) { break }
    } else {
      Write-Host "Имя пользователя и пароль успешно проверены." -f green
      ## Сохраняем учётные данные в переменную в области (scope) Script
      $adcreds = New-Object System.Management.Automation.PSCredential ($($ADCredentials.UserName -replace ".*\\","" -replace "\@.*",""), $ADCredentials.Password)
    }
  }
  return $adcreds
}

Function New-ADComputer-ADSI {
  <#
    .Synopsis
       Creates new computer object in specified AD organization unit.
    .DESCRIPTION
       The function would create computer object in specified AD OU.
       Also if AddUserName is specified, then all priviligies on created computer object would be granted to specified user.
    .EXAMPLE
       New-ADComputer-ADSI -Path "LDAP://OU=Servers,DC=alpha,DC=beta,DC=ru" -Credential $AD_creds -AddUserName "Domain_User" -Name "Server 1"

       Would create computer object "Server 1" in OU Servers of domain alpha.beta.ru with credential specified in PSCredential object $AD_Creds and would grant all priviliges on object to user Domain_User.
    .EXAMPLE
       New-ADComputer-ADSI -Path "LDAP://OU=Servers,DC=alpha,DC=beta,DC=ru" -Name "Server 1"

       Would create computer object "Server 1" in OU Servers of domain alpha.beta.ru with credential of current user.
    .PARAMETER Name
      Name of computer object.
    .PARAMETER Path
      Path to Organization Unit where object should be created.
    .PARAMETER Credential
      Credential to use for object creation.
    .PARAMETER AddUserName
      User to grant all privileges on computer object.
    .NOTES
       Note: you DON'T need RSAT installed for this function to work.

       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Creates new computer object in specified AD organization unit.
  #>
  param(
    [parameter(Mandatory = $true)][string]$Name,
    [parameter(Mandatory = $true)][string]$Path,
    [PSCredential]$Credential,
    [string]$AddUserName
  )
  if ($Credential) {
    $domain_entry = New-Object System.DirectoryServices.DirectoryEntry($Path, $Credential.UserName, $Credential.GetNetworkCredential().Password)
  } else {
    $domain_entry = New-Object System.DirectoryServices.DirectoryEntry($Path)
  }
  
  $ComputerName = $Name
  $samname = $ComputerName + "$"
  # ищем в указанном контейрере AD учётную запись компьютера с заданным именем
  $searcher = [adsisearcher]$Path
  $searcher.Filter = "name=$ComputerName"
  $Computer = [ADSI]$($searcher.FindOne()).Path
  # если учётная запись с таким именем не существует
  if (!($Computer)) {
    # то создаём новую
    $Computer = $domain_entry.create(“Computer”,”cn=$ComputerName”)
    # добавляем минимальный набор атрибутов
    $Computer.put(“sAMAccountName”,$samname) 
    $Computer.put(“userAccountControl”,4128)  
    write-host "Добавляем $ComputerName в нужный контейнер" 
    $Computer.setinfo()
  }     

  $Computer_guid = $newguid = new-object -TypeName System.Guid -ArgumentList (,$Computer.objectGUID[0])
  if ($AddUserName) {
    # Поиск пользователя
    $searcher = [adsisearcher]""
    $searcher.Filter = "userPrincipalName=$AddUserName*"
    $userName = [ADSI]$($searcher.FindOne()).Path
    $User_sid = new-object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList ($userName.objectSid[0],0)

    # Добавление прав
    write-host "Добавляем пользователю $AddUserName полные права на $ComputerName" 
    $ace = new-object System.DirectoryServices.ActiveDirectoryAccessRule $User_sid,"GenericAll","Allow"
    $Computer.get_objectSecurity().AddAccessRule($ace)
  }
  # сохраняем изменения
  $Computer.CommitChanges()
  # Capture any errors (e.g. object already exists) and move on 
  trap 
  { 
    write-host "Error: $_" 
    continue 
  } 
}

Export-ModuleMember -Function Remove-File
Export-ModuleMember -Function Convert-DecToSysNum
Export-ModuleMember -Function Import-ExcelAsCsv
Export-ModuleMember -Function Create-ExcelOfCSV
Export-ModuleMember -Function Get-ADCredentials
Export-ModuleMember -Function New-ADComputer-ADSI