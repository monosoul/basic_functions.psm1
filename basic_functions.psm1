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
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true)][string]$fileName
  )
  if(Test-Path -path $fileName) { Remove-Item -Recurse -path $fileName -Force -Confirm:$false }
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
    .LINK
       https://github.com/monosoul/basic_functions.psm1
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
    .LINK
       https://github.com/monosoul/basic_functions.psm1
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
    .PARAMETER rmlinesbefore
      How many lines should be removed before inserted lines (only for cases when worksheet already exist).
    .NOTES
       You would need MS Office installed for this function to work.

       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Creates one Excel workbook out of one or multiple csv files.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true)][PSObject]$fileslist,
    [parameter(Mandatory = $true)][string]$resultname,
    [int]$rmlinesbefore
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

        $rows_before_import = $worksheetObject.UsedRange.Rows.Count

        ## указываем формат
        $query.TextFileParseType = 1
        $query.TextFileColumnDataTypes = ,1 * $worksheetObject.Cells.Columns.Count
        $query.AdjustColumnWidth = 1
        $query.Refresh() | Out-Null
        $query.Delete() | Out-Null
        $worksheetObject.Rows.Item($($rows_before_import + 1)).Delete() | Out-Null
        if ($rmlinesbefore) {
          1..$rmlinesbefore | %{
            $worksheetObject.Rows.Item($($rows_before_import + 1 - $_)).Delete() | Out-Null
          }
        }
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
    .LINK
       https://github.com/monosoul/basic_functions.psm1
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
      Write-Host "Invalid username or password!!!" -f red
      $rep = Read-Host "Retry? (Y/N)"
      if (($rep -eq "N") -or ($rep -eq "n")) { break }
    } else {
      Write-Host "Username and password are valid." -f green
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
    .LINK
       https://github.com/monosoul/basic_functions.psm1
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
  $searcher.Filter = "(&(name=$ComputerName)(objectClass=computer))"
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

  if ($AddUserName) {
    # Поиск пользователя
    $searcher = [adsisearcher]""
    $searcher.Filter = "(&(|(userPrincipalName=$AddUserName*)(cn=$AddUserName*))(objectClass=User))"
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

Function Extract-ZIPFile {
    <#
    .Synopsis
       Extracts zip archive to selected folder.
    .DESCRIPTION
       Function would extract zip archive to selected folder.
    .EXAMPLE
       Extract-ZIPFile -file "$directorypath\update.zip" -destination "c:\temp"
    .PARAMETER file
      Path to zip archive.
    .PARAMETER destination
      Path to destination folder.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Extracts zip archive to selected folder.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true)][string]$file,
    [parameter(Mandatory = $true)][string]$destination
  )

  Add-Type -AssemblyName "system.io.compression.filesystem"
  [System.IO.Compression.ZipFile]::ExtractToDirectory($file, $destination)
}

Function New-ZIPFile {
    <#
    .Synopsis
       Creates zip archive of selected folder or file.
    .DESCRIPTION
       Function would create zip archive of selected folder or file.
    .EXAMPLE
       New-ZIPFile -source "c:\temp\1" -file "$directorypath\1.zip"
    .PARAMETER file
      Path to zip archive.
    .PARAMETER source
      Path to source folder or file.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Creates zip archive of selected folder or file.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true)][string]$file,
    [parameter(Mandatory = $true)][System.IO.FileInfo]$source
  )

  Add-Type -AssemblyName "system.io.compression.filesystem"
  if ($source.Attributes -notmatch "Directory") {
    $tmpfolder = "$($env:TEMP)\$(([GUID]::NewGuid()).Guid)"
    New-Item -ItemType directory -Name $source.BaseName -Path $tmpfolder -Force -Confirm:$false | Out-Null
    Copy-Item -Path $source.FullName -Destination "$tmpfolder\$($source.BaseName)" -Force -Confirm:$false
    $source = "$tmpfolder\$($source.BaseName)"
  }
  [System.IO.Compression.ZipFile]::CreateFromDirectory($source, $file)
  if ($tmpfolder) {
    Remove-File -fileName $tmpfolder
  }
}

Function ConvertFrom-Json2 {
  <#
    .Synopsis
       Converts a JSON-formatted string to a custom object.
    .DESCRIPTION
       The ConvertFrom-Json2 cmdlet converts a JSON-formatted string to a custom object (PSCustomObject) that has a property for each field in the JSON string. JSON is commonly used by web sites to provide a textual representation of objects.
    
       To generate a JSON string from any object, use the ConvertTo-Json cmdlet.
    .EXAMPLE
       Get-Date | Select-Object -Property * | ConvertTo-Json | ConvertFrom-Json2
    .EXAMPLE
       Invoke-RestMethod -URI "http://somewebservice/method" | ConvertFrom-Json2
    .PARAMETER InputObject
      Specifies the JSON strings to convert to JSON objects. Enter a variable that contains the string, or type a command or expression that gets the string. You can also pipe a string to ConvertFrom-Json2.
        
      The InputObject parameter is required, but its value can be an empty string. When the input object is an empty string, ConvertFrom-Json2 does not generate any output. The InputObject value cannot be null ($null).
    .PARAMETER MaxJsonLength
      Specifies maximum length of JSON-formatted string that function can convert.
    .PARAMETER RecursionLimit
      Specifies how deep would ConvertFrom-Json2 look into JSON objects tree.
    .PARAMETER parallel
      Enable parallel parsing of JSON-formatted string. Parallel parsing is faster, but little more resource-consuming
    .NOTES
       Author:  Andrey Nevedomskiy

       The ConvertFrom-Json2 cmdlet is implemented by using the JavaScriptSerializer class (http://msdn.microsoft.com/en-us/library/system.web.script.serialization.javascriptserializer(VS.100).aspx).
    .FUNCTIONALITY
       Converts a JSON-formatted string to a custom object.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline=$True, Position = 0)]
    [String]$InputObject,
    [Parameter(Mandatory = $false, Position = 1)]
    [long]$MaxJsonLength = 9999999,
    [Parameter(Mandatory = $false, Position = 2)]
    [long]$RecursionLimit = 100,
    [Parameter(Mandatory = $false)]
    [switch]$parallel
  )
  $ChildFunctions = {
    Function ParseItem {
      param(
        $jsonItem
      )
    
      if($jsonItem.PSObject.TypeNames -match "Array") {
        return ParseJsonArray($jsonItem)
      }
      elseif($jsonItem.PSObject.TypeNames -match "Hashtable") {
        return ParseJsonObject($jsonItem)
      }
      else {
        return $jsonItem
      }
    }
    
    Function ParseJsonObject {
      param(
        $jsonObj
      )
    
      $result = New-Object -TypeName PSCustomObject
      foreach ($key in $jsonObj.Keys) {
        Remove-Variable -Name item -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Remove-Variable -Name parsedItem -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $item = $jsonObj[$key]
        if ($item) {
          $parsedItem = ParseItem $item
        }
        else {
          $parsedItem = $null
        }
        $result | Add-Member -MemberType NoteProperty -Name $key -Value $parsedItem
      }
      return $result
    }
    
    Function ParseJsonArray {
      param(
        $jsonArray
      )
    
      $result = @()
      $jsonArray | ForEach-Object {
        $result += , (ParseItem $_)
      }
      return $result
    }
  }

  [System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions") | Out-Null
  $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = $MaxJsonLength
  $ser.RecursionLimit = $RecursionLimit

  [System.Collections.Hashtable[]]$InputObject = $ser.DeserializeObject($InputObject)
  $InputObject = {$InputObject}.Invoke()
  [System.GC]::Collect()

  if ($parallel.IsPresent) {
    [int]$cpunum = 0
    (Get-WmiObject -Class Win32_processor).NumberOfLogicalProcessors | ForEach-Object {
      $cpunum += $_
    }
    $throttle = $cpunum * 2

    # getting amount of hashes per thread
    $limit = [math]::Ceiling(($InputObject.Count / $throttle))

    # grouping hashes by threads
    $groups = @()
    [int]$skip = 0
    1..$throttle | %{
      $groups += New-Object PSObject -Property @{
        Name = $_
        Value = ($InputObject | Select -Skip $skip -First $limit)
      }
      $skip += $limit
    }

    # parallelizing output parsing
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $throttle)
    $RunspacePool.Open()

    [string]$RS_scriptblock = "
      param(
        [PSObject]`$hashes,
        [ScriptBlock]`$ChildFunctions
      )
      
    "
    $RS_scriptblock += $ChildFunctions.ToString()
    $RS_scriptblock += "
      Invoke-Command -NoNewScope -ScriptBlock `$ChildFunctions

      return ParseItem(`$hashes)
    "
    [ScriptBlock]$RS_scriptblock = [System.Management.Automation.ScriptBlock]::Create($RS_scriptblock)

    $Jobs = New-Object System.Collections.ArrayList

    ForEach ($group in $groups) {
      $Job = [powershell]::Create().AddScript($RS_scriptblock)

      $Job.AddArgument($group.Value) | Out-Null
      $Job.AddArgument($ChildFunctions) | Out-Null

      $Job.RunspacePool = $RunspacePool
      $Jobs.Add((New-Object PSObject -Property @{
        Pipe = $Job
        Result = $Job.BeginInvoke()
      })) | Out-Null
    }

    #Waiting for all jobs to end
    $counter = 0
    $jobs_count = @($Jobs).Count
    $data = @()
    Do {
      #saving results
      ForEach ($Job in $Jobs) {
        if ($Job.Result.IsCompleted) {
          Remove-Variable -Name tmpRes -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
          $tmpRes = $Job.Pipe.EndInvoke($Job.Result)
          $data += &{if ($tmpRes) {$tmpRes}}
          $Job.Pipe.dispose()
          $Job.Result = $null
          $Job.Pipe = $null
          $counter++
        }
      }
      #removing unused jobs (runspaces)
      $temphash = $Jobs.Clone()
      $temphash | Where-Object { $_.pipe -eq $null } | ForEach {
        $Jobs.Remove($_) | Out-Null
      }
      Remove-Variable -Name tmpRes -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
      Remove-Variable -Name temphash -Force -Confirm:$false
      [System.GC]::Collect()
      Start-Sleep -Seconds 1
    } While ( $counter -lt $jobs_count )

    $RunspacePool.dispose()
    $RunspacePool.Close()
    Remove-Variable -Name Job -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name Jobs -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name RunspacePool -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name RS_scriptblock -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name group -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name groups -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name limit -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Remove-Variable -Name skip -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    [System.GC]::Collect()
    return $data
  }
  else {
    Invoke-Command -NoNewScope -ScriptBlock $ChildFunctions
    return ParseItem($InputObject)
  }
}

Function Get-SqlServerData {
  <#
    .Synopsis
       This function connects to specified SQL Server DB and returns queried data.
    .DESCRIPTION
       This function connects to specified SQL Server DB and returns queried data.
    .EXAMPLE
       $result = Get-DatabaseData -isSQLServer -query "Select * from users" -connectionString "Data Source=127.0.0.1;Initial Catalog=SampleDB;Trusted_Connection=True;"
    .EXAMPLE
       $result = Get-DatabaseData -isSQLServer -query "Select * from users" -connectionString "Data Source=127.0.0.1;Initial Catalog=SampleDB;User Id=myLogin;Password=pwd;"
    .PARAMETER query
      Select query to execute.
    .PARAMETER connectionString
      Connection string.
    .NOTES
       Author:  Don Jones
    .FUNCTIONALITY
       This function connects to specified SQL Server DB and returns queried data.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $true)]
    [string]$connectionString,
    [parameter(Mandatory = $true)][ValidatePattern("^select.*")]
    [string]$query,
    [switch]$isSQLServer
  )
  if ($isSQLServer) {
    Write-Verbose 'in SQL Server mode'
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
  } else {
    Write-Verbose 'in OleDB mode'
    $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
  }
  $connection.ConnectionString = $connectionString
  $command = $connection.CreateCommand()
  $command.CommandText = $query
  if ($isSQLServer) {
    $adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
  } else {
    $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
  }
  $dataset = New-Object -TypeName System.Data.DataSet
  $adapter.Fill($dataset) | Out-Null
  $dataset.Tables[0]
}

Function Invoke-SqlServerQuery {
  <#
    .Synopsis
       This function connects to specified SQL Server DB and executes any provided query.
    .DESCRIPTION
       This function connects to specified SQL Server DB and executes any provided query.
    .EXAMPLE
       $result = Get-DatabaseData -isSQLServer -query "delete from users" -connectionString "Data Source=127.0.0.1;Initial Catalog=SampleDB;Trusted_Connection=True;"
    .EXAMPLE
       $result = Get-DatabaseData -isSQLServer -query "delete from users" -connectionString "Data Source=127.0.0.1;Initial Catalog=SampleDB;User Id=myLogin;Password=pwd;"
    .PARAMETER query
      Select query to execute.
    .PARAMETER connectionString
      Connection string.
    .NOTES
       Author:  Don Jones
    .FUNCTIONALITY
       This function connects to specified SQL Server DB and executes any provided query.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $true)]
    [string]$connectionString,
    [parameter(Mandatory = $true)]
    [string]$query,
    [switch]$isSQLServer
  )
  if ($isSQLServer) {
    Write-Verbose 'in SQL Server mode'
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
  } else {
    Write-Verbose 'in OleDB mode'
    $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
  }
  $connection.ConnectionString = $connectionString
  $command = $connection.CreateCommand()
  $command.CommandText = $query
  $connection.Open()
  $command.ExecuteNonQuery()
  $connection.close()
}

Function Convert-BashScriptToOneLiner {
  <#
    .Synopsis
       This function converts multi-lined bash script to one-liner.
    .DESCRIPTION
       This function converts multi-lined bash script to one-liner.
    .EXAMPLE
       $oneliner = Convert-BashScriptToOneLiner -filePath="C:\Data\Work\Scripts\Get-Solaris-metrics\cperf.sh" -ForInvokeVMScript

       $vm | Invoke-VMScript -ScriptText $oneliner
    .PARAMETER contents
      Contents of script to convert.
    .PARAMETER filePath
      Path to script to convert.
    .PARAMETER ForInvokeVMScript
      Additional symbol escapes for usage with Invoke-VMSCript cmdlet.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       This function converts multi-lined bash script to one-liner.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [Parameter(ParameterSetName = "AsString", Mandatory = $true)]
    [PSObject]$contents,
    [Parameter(ParameterSetName = "AsPath", Mandatory = $true)]
    [string]$filePath,
    [Parameter(ParameterSetName = "AsPath", Mandatory = $false)]
    [Parameter(ParameterSetName = "AsString", Mandatory = $false)]
    [switch]$ForInvokeVMScript
  )

  if ($filePath) {
    $contents = Get-Content -Encoding Default -Path $filePath
  }

  [string[]]$contents = [string[]]$contents

  #удаляем строки, в которых только комментарии
  $contents = $contents | Where-Object { $_ -notmatch "^([ ]*|[\t]*)#" }

  #удаляем комментарии
  $contents = $contents -replace "[^\\]#.*$",""

  #удаляем пустые строки
  $contents = $contents | Where-Object { $_ -match "[A-Za-z0-9()\[\]\{\}]+" }

  if ($ForInvokeVMScript.IsPresent) {
    # экранируем символ ` в командах
    $contents = $contents -replace "``","\\\``"
  }

  #выделяем then, do, else и {
  $contents = $contents -replace "([ ]+|[\t]+|^)(do|then|else|\{)([ ]+|[\t]+|$)","`$1`$2;`$3"

  #добавляем пробелы в начало каждой строки
  $contents = $contents -replace "^"," "

  #объединяем строки
  $contents = $contents -join ";"

  #удаляем ; после then, do, else и {
  $contents = $contents -replace "(then|do|else|\{);[ ]*[\t]*;","`$1"

  #удаляем лишний пробел в начале
  $contents = $contents -replace "^([ ]+|[\t]+)"

  return $contents
}

Function Escape-SpecialCharacters {
  <#
    .Synopsis
       This function escapes special characters in input string.
    .DESCRIPTION
       This function escapes special characters in input string.
    .EXAMPLE
       C:\PS>$var = @()
       C:\PS>$var += "asdsadaweq1 ,.3v12415][ad"
       C:\PS>$var += "./.1v32\1v4]2154"

       C:\PS>$var | Escape-SpecialCharacters

       asdsadaweq1 ,\.3v12415\]\[ad
       \./\.1v32\\1v4\]2154
    .PARAMETER string
      Input string.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       This function escapes special characters in input string.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline=$True, Position = 0)]
    [string[]]$string
  )

  process{
    $result = @()
    ForEach($s in $string) {
      $result += $s -replace "([\[\]\\\|\.\*\?\+\(\)\{\}])","\`$1"
    }
    return $result
  }
}

Function Start-RSJobs {
  <#
    .Synopsis
       Starts a Windows PowerShell background RunSpace jobs.
    .DESCRIPTION
       The Start-RSJob cmdlet starts a Windows PowerShell background RunSpace jobs on the local computer.

       A Windows PowerShell background RunSpace jobs runs a command(s) "in the background" without interacting with the current session. When you start a background RunSpace jobs, a job objects is returned immediately, even if the job takes an extended time to complete. You can continue to work in the session without interruption while the job runs.
    .EXAMPLE
       Start-RSJobs -ScriptBlock { Start-Sleep -Seconds 5 }

       Would run lCPU*2 jobs which would be sleeping for 5 seconds.
    .EXAMPLE
       C:\PS>$var = [System.Collections.Queue]::Synchronized( (1..10) )
       C:\PS>$result = [System.Collections.ArrayList]::Synchronized( (New-Object System.Collections.ArrayList) )

       C:\PS>Start-RSJobs -ScriptBlock {
                                         param(
                                           $var,
                                           $result
                                         )

                                         while($var.Count -gt 0) {
                                           $result.Add($var.Dequeue() * (Get-Random))
                                         }
                                       } `
                          -throttle 4 `
                          -ArgumentList $var, $result

       Would run 4 RS jobs which would multiply every number from $var queue, multiply it by random number and write result to $result.
    .EXAMPLE
       C:\PS>$var = [System.Collections.Queue]::Synchronized( (1..10) )
       C:\PS>$result = [System.Collections.ArrayList]::Synchronized( (New-Object System.Collections.ArrayList) )

       C:\PS>Start-RSJobs -ScriptBlock {
                                         param(
                                           $var,
                                           $result
                                         )

                                         while($var.Count -gt 0) {
                                           $result.Add($var.Dequeue() * (Get-Random))
                                         }
                                       } `
                          -throttle 4 `
                          -ArgumentList $var, $result `
             | Wait-RSJob | Remove-RSJob

       Basically result would be as in previous example, but it'll also wait for RS jobs to end and after that removes them.
    .EXAMPLE
       C:\PS>$ScriptBlocks = @()
       C:\PS>$Arguments = New-Object System.Collections.ArrayList
       C:\PS>1..10 | %{
               $ScriptBlocks += [ScriptBlock]::Create("return ($_ + `$args[0] + `$args[1])")
               $Arguments.Add(@((Get-Random), (Get-Random))) | Out-Null
             }

       C:\PS>Start-RSJobs -ScriptBlock $ScriptBlocks -modulesToImport $modules -snapinsToImport $snapins -throttle 4 -ArgumentList $Arguments | Wait-RSJob | Receive-RSJob


       3720834745
       1969195663
       1142950846
       1936216172
       3518711541
       1869258939
       2082920156
       1820351727
       2001528136
       3318067845


       In this example would be created 10 runspaces with 10 different script blocks and 10 different arguments lists. Runspaces would be throttled by 4 runspaces at a time.
       Each argument list would contain 2 random numbers.
    .PARAMETER ScriptBlock
      Specifies the commands to run in the background job. Enclose the commands in braces ( { } ) to create a script block. This parameter is required.
    .PARAMETER modulesToImport
      Specifies the modules to import in the background job.
    .PARAMETER snapinsToImport
      Specifies the snap-ins to import in the background job.
    .PARAMETER functionsToImport
      Specifies the functions to import in the background job.
    .PARAMETER functionsToImport
      Specifies the functions to import in the background job.
    .PARAMETER variablesToImport
      Specifies the variables to import in the background job.
    .PARAMETER ArgumentList
      Specifies the arguments (parameter values) for the script that is specified by the ScriptBlock parameter.
        
      Because all of the values that follow the ArgumentList parameter name are interpreted as being values of ArgumentList, the ArgumentList parameter should be the last parameter in the command.
    .PARAMETER throttle
      Specifies amount of threads.
    .PARAMETER CurrentHostOutput
      Specifies whether should RunSpace pool use current host for output or not.
    .PARAMETER ApartmentState
      Specifies ApartementState for RunSpace pool.
      DO NOT TOUCH THIS PARAMETER IF YOU DON'T KNOW WHAT YOU'RE DOING!
      
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       The Start-RSJob cmdlet starts a Windows PowerShell background RunSpace jobs on the local computer.
    .OUTPUTS
       PSCustomObject containing RS Jobs
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$True)]
    [ScriptBlock[]]$ScriptBlock,
    [parameter(Mandatory = $false)]
    [string[]]$modulesToImport,
    [parameter(Mandatory = $false)]
    [string[]]$snapinsToImport,
    [parameter(Mandatory = $false)]
    [string[]]$functionsToImport,
    [parameter(Mandatory = $false)]
    [string[]]$variablesToImport,
    [parameter(Mandatory = $false)]
    [Object[]]$ArgumentList,
    [parameter(Mandatory = $false)]
    [int]$throttle=$(&{
      # Default throttle value is number of logical CPUs multiplied by 2, i.e. 2 threads per lCPU
      [int]$cpunum = 0
      (Get-WmiObject -Class Win32_processor).NumberOfLogicalProcessors | ForEach-Object {
        $cpunum += $_
      }
      $cpunum * 2
    }),
    [parameter(Mandatory = $false)]
    [switch]$CurrentHostOutput,
    [parameter(Mandatory = $false)]
    [System.Threading.ApartmentState]$ApartmentState = [System.Threading.ApartmentState]::Unknown
  )

  Begin{

    if ($Global:RunSpacePools -eq $null) {
      $Global:RunSpacePools = [System.Collections.ArrayList]::Synchronized( (New-Object System.Collections.ArrayList) )
    }
    if ($Global:RSJobs -eq $null) {
      $Global:RSJobs = [System.Collections.ArrayList]::Synchronized( (New-Object System.Collections.ArrayList) )
    }
    if ($Global:RSJobID -eq $null) {
      [int64]$Global:RSJobID = 0
    }

    Write-Verbose "Initializaing initial session state for RunSpace pool"
    $initialSessionState = [InitialSessionState]::CreateDefault()

    #setting apartment state
    $initialSessionState.ApartmentState = $ApartmentState

    #importing snap-ins
    ForEach ($snapIn in $snapinsToImport) {
      Write-Verbose "Trying to import snap-in: $snapIn"
      try{
        $initialSessionState.ImportPSSnapIn($snapIn,[ref]'') | Out-Null
      }
      catch{
        Write-Warning "Wasn't able to import Snap-in: $snapIn`r`nError: $_"
      }
    }

    #importing modules
    ForEach ($module in $modulesToImport) {
      Write-Verbose "Trying to import module: $module"
      $initialSessionState.ImportPSModule($module)
    }

    #importing functions
    ForEach ($function in $functionsToImport) {
      Write-Verbose "Trying to import function $function"
      try{
        $definition = Get-Content Function:\$function -ErrorAction Stop
        Write-Verbose "Function $function definition:`r`n$definition"
        $sessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $function, $definition
        $initialSessionState.Commands.Add($sessionStateFunction)
      }
      catch{
        Write-Warning "Wasn't able to import function: $function`r`nError: $_"
      }
    }

    #importing variables
    ForEach ($variable in $variablesToImport) {
      Write-Verbose "Trying to import variable $variable"
      try{
        $value = (Get-Variable -Name $variable -ErrorAction Stop).Value
        Write-Verbose "Function $variable value:`r`n$value"
        $sessionStateVariable = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $variable, $value, $variable
        $initialSessionState.Variables.Add($sessionStateVariable)
        Remove-Variable -Name value -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
      }
      catch{
        Write-Warning "Wasn't able to import variable: $variable`r`nError: $_"
      }
    }

    Write-Verbose "Creating RunSpace pool for $throttle runspaces"
    if ($CurrentHostOutput.IsPresent) {
      $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $throttle, $initialSessionState, $Host)
    }
    else {
      $RunspacePool = [RunspaceFactory]::CreateRunspacePool($initialSessionState)
      $RunspacePool.SetMinRunspaces(1) | Out-Null
      $RunspacePool.SetMaxRunspaces($throttle) | Out-Null
    }
    $RunspacePool | Add-Member -NotePropertyName "Runspaces" -NotePropertyValue 0
    $Global:RunSpacePools.Add($RunspacePool) | Out-Null
    $RunspacePool.Open()

  }

  Process{
    if ($ScriptBlock.Count -eq 1) {
      $RSamount = $throttle
    }
    else {
      $RSamount = $ScriptBlock.Count
      [Object[][]]$ArgumentList = [Object[][]]$ArgumentList
    }

    $Private:RSJobs = New-Object System.Collections.ArrayList

    $RScounter = 0
    1..$RSamount | %{
      Write-Verbose "Creating runspace #$_"

      Write-Verbose "Adding arguments to runspace"
      if ($ScriptBlock.Count -eq 1) {
        $Job = [powershell]::Create().AddScript($ScriptBlock, $true)
        ForEach ($argument in $ArgumentList) {
          $Job.AddArgument($argument) | Out-Null
        }
      }
      else {
        $Job = [powershell]::Create().AddScript($ScriptBlock[$RScounter], $true)
        ForEach ($argument in $ArgumentList[$RScounter]) {
          $Job.AddArgument($argument) | Out-Null
        }
      }

      Write-Verbose "Assigning runspace to pool"
      $Job.RunspacePool = $RunspacePool
      $Job.RunspacePool.Runspaces++
      $Private:RSJobs.Add((New-Object PSObject -Property @{
        Pipe = $Job
        Result = $Job.BeginInvoke()
        Id = ++$Global:RSJobID
      })) | Out-Null
      $RScounter++
    }

    $Private:RSJobs | %{
      $Global:RSJobs.Add($_) | Out-Null
    }

    return $Private:RSJobs
  }
}

Function Get-RSJob {
  <#
    .Synopsis
       Gets Windows PowerShell background RunSpace jobs that are running in the current session
    .DESCRIPTION
       The Get-RSJob cmdlet gets objects that represent the background RunSpace jobs that were started in the current session. You can use Get-Job to get RunSpace jobs that were started by using the Start-RSJobs cmdlet.
    
       Without parameters, a "Get-RSJob" command gets all RunSpace jobs in the current session. You can use the parameters of Get-RSJob to get particular jobs.
    
       The job object that Get-RSJob returns contains useful information about the job, but it does not contain the job results. To get the results, use the Receive-RSJob cmdlet.
    
       A Windows PowerShell background RunSpace job is a command that runs "in the background" without interacting with the current session. Typically, you use a background RunSpace job to run a complex command that takes a long time to complete.
    .EXAMPLE
       Get-Job

       This command gets all background RunSpace jobs started in the current session. It does not include jobs created in other sessions, even if the jobs run on the local computer.
    .EXAMPLE
       Get-Job -Id 1

       This command gets background RunSpace job with ID 1.
    .PARAMETER id
       Gets only RunSpace jobs with the specified IDs.
        
       The ID is an integer that uniquely identifies the RunSpace job within the current session. It is easier to remember and to type than the instance ID, but it is unique only within the current session. You can type one or more IDs (separated by commas). To find the ID of a RunSpace job, type "Get-RSJob" without parameters.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Gets Windows PowerShell background RunSpace jobs that are running in the current session
    .OUTPUTS
       PSCustomObject containing RS Jobs
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $false)]
    [int64[]]$id
  )

  if ($id) {
    $result = @()
    ForEach ($i in $id) {
      $result += ($Global:RSJobs | ?{ $_.Id -eq $i })
    }
    return $result
  }
  else {
    return $Global:RSJobs
  }
}

Function Receive-RSJob {
  <#
    .Synopsis
       Gets the results of the Windows PowerShell background RunSpace jobs in the current session.
    .DESCRIPTION
       The Receive-RSJob cmdlet gets the results of Windows PowerShell background RunSpace jobs, such as those started by using the Start-RSJobs cmdlet. You can get the results of all jobs or identify jobs by their ID or by submitting a job object.
    
       When you start a Windows PowerShell background RunSpace job, the job starts, but the results do not appear immediately. Instead, the command returns an object that represents the background RunSpace job. The job object contains useful information about the job, but it does not contain the results. This method allows you to continue working in the session while the job runs.
    
       The Receive-RSJob cmdlet gets the results of only completed jobs. You'll need to stop the job or to wait until the job ends before receiving results.
    .EXAMPLE
       PS C:\>$job = Start-RSJobs -ScriptBlock {Get-Process}
       PS C:\>Receive-RSJob -Job $job

       These commands use the Job parameter of Receive-RSJob to get the results of a particular RunSpace job.
    
       The first command uses the Start-RSJobs cmdlet to start a RunSpace job that runs a Get-Process command. The command uses the assignment operator (=) to save the resulting job object in the $job variable.
    
       The second command uses the Receive-RSJob cmdlet to get the results of the RunSpace job. It uses the Job parameter to specify the job.
    .EXAMPLE
       PS C:\>Get-RSJob | Receive-RSJob
    .PARAMETER id
       Gets the results of RunSpace jobs with the specified IDs.
        
       The ID is an integer that uniquely identifies the job within the current session. It is easier to remember and type than the instance ID, but it is unique only within the current session. You can type one or more IDs (separated by commas). To find the ID of a job, type "Get-RSJob" without parameters.
    .PARAMETER job
       Specifies the RunSpace job for which results are being retrieved. This parameter is required in a Receive-RSJob command. Enter a variable that contains the job or a command that gets the job. You can also pipe a job object to Receive-RSJob.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Gets the results of the Windows PowerShell background RunSpace jobs in the current session.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true, ParameterSetName = "id")]
    [int64[]]$id,
    [parameter(Mandatory = $true, ParameterSetName = "job", ValueFromPipeline=$True)][ValidateScript({
      ForEach ($j in $_) {
        if ($j.GetType().Name -ne "PSCustomObject") {
          return $false
        }
        if (@($j.psobject.Properties).Count -eq 3) {
          ForEach ($property in $j.psobject.Properties) {
            if (($property.Name -ne "Id") -and ($property.Name -ne "Pipe") -and ($property.Name -ne "Result")) {
              return $false
            }
          }
          return $true
        } else {
          return $false
        }
      }
    })]
    [PSCustomObject[]]$job
  )

  Begin{
    if ($id) {
      [PSCustomObject[]]$job = Get-RSJob -id $id
    }

    $data = @()
  }

  process{

    ForEach ($j in $job) {
      if ($j.Result.IsCompleted) {
        $data += $j.Pipe.EndInvoke($j.Result)
      } else {
        Write-Warning "Job $($j.id) isn't completed yet!"
      }
    }

  }

  End{
    return $data
  }
}

Function Stop-RSJob {
  <#
    .Synopsis
       Stops a Windows PowerShell background RunSpace job.
    .DESCRIPTION
       The Stop-RSJob cmdlet stops Windows PowerShell background RunSpace jobs that are in progress. You can use this cmdlet to stop selected jobs based on their ID or by passing a job object to Stop-RSJob.
    
       You can use Stop-RSJob to stop background RunSpace jobs, such as those that were started by using the Start-RSJobs cmdlet. When you stop a background RunSpace job, Windows PowerShell completes all tasks that are pending in that job queue and then ends the job. No new tasks are added to the queue after this command is submitted.
    
       This cmdlet does not delete background RunSpace jobs. To delete a job, use the Remove-RSJob cmdlet.
    .EXAMPLE
       PS C:\>Stop-RSJob -ID 1, 3, 4
       
       This command stops three jobs. It identifies them by their IDs
    .EXAMPLE
       Get-RSJob | Stop-RSJob

       This command stops all current RunSpace jobs.
    .PARAMETER id
       Stops RunSpace jobs with the specified IDs.
        
       The ID is an integer that uniquely identifies the RunSpace job within the current session. It is easier to remember and type than the InstanceId, but it is unique only within the current session. You can type one or more IDs (separated by commas). To find the ID of a job, type "Get-RSJob" without parameters.
    .PARAMETER job
       Specifies the RunSpace jobs to be stopped. Enter a variable that contains the jobs or a command that gets the jobs. You can also use a pipeline operator to submit jobs to the Stop-RSJob cmdlet.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Stops a Windows PowerShell background RunSpace job.
    .OUTPUTS
       PSCustomObject containing RS Jobs
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true, ParameterSetName = "id")]
    [int64[]]$id,
    [parameter(Mandatory = $true, ParameterSetName = "job", ValueFromPipeline=$True)][ValidateScript({
      ForEach ($j in $_) {
        if ($j.GetType().Name -ne "PSCustomObject") {
          return $false
        }
        if (@($j.psobject.Properties).Count -eq 3) {
          ForEach ($property in $j.psobject.Properties) {
            if (($property.Name -ne "Id") -and ($property.Name -ne "Pipe") -and ($property.Name -ne "Result")) {
              return $false
            }
          }
          return $true
        } else {
          return $false
        }
      }
    })]
    [PSCustomObject[]]$job
  )

  Begin{
    if ($id) {
      [PSCustomObject[]]$job = Get-RSJob -id $id
    }
  }

  process {

    ForEach ($j in $job) {
      try{
        $j.Pipe.Stop()
      }
      catch{
        Write-Warning "Wasn't able to stop the job $($j.id)! Error: $_"
      }
    }

    return $job
  }
}

Function Remove-RSJob {
  <#
    .Synopsis
       Deletes a Windows PowerShell background RunSpace job.
    .DESCRIPTION
       The Remove-RSJob cmdlet deletes Windows PowerShell background RunSpace jobs that were started by using the Start-RSJobs.
    
       You can use this cmdlet to delete all RunSpace jobs or delete selected jobs based on their ID or by passing a job object to Remove-RSJob.
    
       Before deleting a running RunSpace job, use the Stop-RSJob cmdlet to stop the job. If you try to delete a running job, the command fails. You can use the Force parameter of Remove-RSJob to delete a running job.
    
       If you do not delete a background RunSpace job, the job remains in the global job cache until you close the session in which the job was created.
    .EXAMPLE
       PS C:\>Remove-RSJob -ID 1, 3, 4
       
       This command deletes three jobs. It identifies them by their IDs
    .EXAMPLE
       Get-RSJob | Remove-RSJob

       This command deletes all current RunSpace jobs.
    .PARAMETER id
       Deletes background RunSpace jobs with the specified IDs.
        
       The ID is an integer that uniquely identifies the RunSpace job within the current session. It is easier to remember and type than the instance ID, but it is unique only within the current session. You can type one or more IDs (separated by commas). To find the ID of a job, type "Get-RSJob" without parameters.
    .PARAMETER job
       Specifies the RunSpace jobs to be deleted. Enter a variable that contains the jobs or a command that gets the jobs. You can also use a pipeline operator to submit jobs to the Remove-RSJob cmdlet.
    .PARAMETER Force
       Deletes the RunSpace job even if the status is "Running". Without the Force parameter, Remove-RSJob does not delete running jobs.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Deletes a Windows PowerShell background RunSpace job.
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true, ParameterSetName = "id")]
    [int64[]]$id,
    [parameter(Mandatory = $true, ParameterSetName = "job", ValueFromPipeline=$True)][ValidateScript({
      ForEach ($j in $_) {
        if ($j.GetType().Name -ne "PSCustomObject") {
          return $false
        }
        if (@($j.psobject.Properties).Count -eq 3) {
          ForEach ($property in $j.psobject.Properties) {
            if (($property.Name -ne "Id") -and ($property.Name -ne "Pipe") -and ($property.Name -ne "Result")) {
              return $false
            }
          }
          return $true
        } else {
          return $false
        }
      }
    })]
    [PSCustomObject[]]$job,
    [parameter(Mandatory = $false, ParameterSetName = "id")]
    [parameter(Mandatory = $false, ParameterSetName = "job")]
    [switch]$Force
  )

  Begin{
    if ($id) {
      [PSCustomObject[]]$job = Get-RSJob -id $id
    }
  }

  process {

    ForEach ($j in $job) {
      if ($j.Result.IsCompleted -or $Force.IsPresent) {
        $RunspacePool = $j.Pipe.RunspacePool
        if (!$j.Result.IsCompleted) {
          $j.Pipe.Stop()
        }
        $j.Pipe.dispose()
        $j.Result = $null
        $j.Pipe = $null
        $RunspacePool.Runspaces--
        if ($RunspacePool.GetMaxRunspaces() -eq $RunspacePool.GetAvailableRunspaces()) {
          if ($RunspacePool.Runspaces -eq 0) {
            $RunspacePool.dispose()
            $RunspacePool.Close()
            $Global:RunSpacePools.Remove($RunspacePool)
          }
        }
        Remove-Variable -Name RunspacePool -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        [System.GC]::Collect()
      }
      else {
        Write-Warning "Job $($j.Id) isn't completed yet!"
      }
    }

  }

  End{
    #removing unused runspaces
    $temphash = $Global:RSJobs.Clone()
    $temphash | Where-Object { $_.pipe -eq $null } | ForEach {
      $Global:RSJobs.Remove($_) | Out-Null
    }
    Remove-Variable -Name temphash -Force -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    [System.GC]::Collect()
  }
}

Function Wait-RSJob {
  <#
    .Synopsis
       Suppresses the command prompt until one or all of the Windows PowerShell background RunSpace jobs running in the session are complete.
    .DESCRIPTION
       The Wait-RSJob cmdlet waits for Windows PowerShell background RunSpace jobs to complete before it displays the command prompt. You can wait until any background RunSpace job is complete, or until all background jobs are complete, and you can set a maximum wait time for the job.
    
       When the commands in the RunSpace job are complete, Wait-RSJob displays the command prompt and returns a job object so that you can pipe it to another command.
    
       You can use Wait-RSJob cmdlet to wait for background RunSpace jobs, such as those that were started by using the Start-RSJobs cmdlet.
    .EXAMPLE
       PS C:\>Wait-RSJob -ID 1, 3, 4
       
       This command waits for three jobs to end. It identifies them by their IDs
    .EXAMPLE
       Get-RSJob | Wait-RSJob

       This command waits for all current RunSpace jobs to end.
    .PARAMETER id
       Waits for RunSpace jobs with the specified IDs.
        
       The ID is an integer that uniquely identifies the RunSpace job within the current session. It is easier to remember and type than the InstanceId, but it is unique only within the current session. You can type one or more IDs (separated by commas). To find the ID of a job, type "Get-RSJob" without parameters.
    .PARAMETER job
       Waits for the specified RunSpace jobs. Enter a variable that contains the job objects or a command that gets the job objects. You can also use a pipeline operator to send job objects to the Wait-RSJob cmdlet.
    .PARAMETER Timeout
       Determines the maximum wait time for each background RunSpace job, in seconds. The default, -1, waits until the job completes, no matter how long it runs. The timing starts when you submit the Wait-RSJob command, not the Start-RSJob command.
        
       If this time is exceeded, the wait ends and the command prompt returns, even if the job is still running. No error message is displayed.
    .NOTES
       Author:  Andrey Nevedomskiy
    .FUNCTIONALITY
       Suppresses the command prompt until one or all of the Windows PowerShell background RunSpace jobs running in the session are complete.
    .OUTPUTS
       PSCustomObject containing RS Jobs
    .LINK
       https://github.com/monosoul/basic_functions.psm1
  #>
  param(
    [parameter(Mandatory = $true, ParameterSetName = "id")]
    [int64[]]$id,
    [parameter(Mandatory = $true, ParameterSetName = "job", ValueFromPipeline=$True)][ValidateScript({
      ForEach ($j in $_) {
        if ($j.GetType().Name -ne "PSCustomObject") {
          return $false
        }
        if (@($j.psobject.Properties).Count -eq 3) {
          ForEach ($property in $j.psobject.Properties) {
            if (($property.Name -ne "Id") -and ($property.Name -ne "Pipe") -and ($property.Name -ne "Result")) {
              return $false
            }
          }
          return $true
        } else {
          return $false
        }
      }
    })]
    [PSCustomObject[]]$job,
    [int]$Timeout=-1
  )

  Begin{
    if ($id) {
      [PSCustomObject[]]$job = Get-RSJob -id $id
    }
    $startTime = Get-Date
  }

  process{
    $counter = 0
    $jobs_count = $job.Count
    $data = @()
    Do {
      ForEach ($j in $job) {
        if ($j.Result.IsCompleted) {
          $counter++
        }
      }
      Start-Sleep -Seconds 1
    } While ( ($counter -lt $jobs_count) -and (([math]::Round(((Get-Date) - $startTime).TotalSeconds) -le $Timeout) -or ($Timeout -eq -1)) )

    return $job
  }
}

Export-ModuleMember -Function Remove-File
Export-ModuleMember -Function Convert-DecToSysNum
Export-ModuleMember -Function Import-ExcelAsCsv
Export-ModuleMember -Function Create-ExcelOfCSV
Export-ModuleMember -Function Get-ADCredentials
Export-ModuleMember -Function New-ADComputer-ADSI
Export-ModuleMember -Function Extract-ZIPFile
Export-ModuleMember -Function New-ZIPFile
Export-ModuleMember -Function ConvertFrom-Json2
Export-ModuleMember -Function Get-SqlServerData
Export-ModuleMember -Function Invoke-SqlServerQuery
Export-ModuleMember -Function Convert-BashScriptToOneLiner
Export-ModuleMember -Function Escape-SpecialCharacters
Export-ModuleMember -Function Start-RSJobs
Export-ModuleMember -Function Get-RSJob
Export-ModuleMember -Function Receive-RSJob
Export-ModuleMember -Function Stop-RSJob
Export-ModuleMember -Function Remove-RSJob
Export-ModuleMember -Function Wait-RSJob
