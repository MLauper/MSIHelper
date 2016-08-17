

$MSIInfo = & "C:\Program Files (x86)\Windows Kits\10\bin\x86\MsiInfo.exe" "C:\Program Files (x86)\Windows Kits\10\bin\x86\Orca-x86_en-us.msi" | Out-String

$MSIInfo | clip 

$MSIInfo = @"

Class Id for the MSI storage is {000C1084-0000-0000-C000-000000000046}

[ 1][/c] Codepage = 1252

[ 2][/t] Title = Installation Database

[ 3][/j] Subject = Orca

[ 4][/a] Author = Microsoft Corporation

[ 5][/k] Keywords = Installer

[ 6][/o] Comments = This installer database contains the logic and data required to install Orca.

[ 7][/p] Template(MSI CPU,LangIDs) = Intel;1033

[ 9][/v] Revision = {6EBA87D8-9142-4C2A-902D-6F3B86340D44}

[12][/r] Created = 2016/07/29 06:18:14

[13][/q] LastSaved = 2016/07/29 06:18:14

[14][/g] Pages(MSI Version Used) = 400

[15][/w] Words(MSI Source and Elevation Prompt Type) = 2

[18][/n] Application = Windows Installer XML (3.7.4128.0)

[19][/u] Security = 2

MsiInfo V 5.0
Copyright (c) Microsoft Corporation. All Rights Reserved
"@

$valueType = @"
0;"PID_DICTIONARY";"Not applicable. Reserved."
1;"PID_CODEPAGE";"Codepage Summary Property."
2;"PID_TITLE";"Title Summary Property."
3;"ID PID_SUBJECT";"Subject Summary Property."
4;"PID_AUTHOR";"Author Summary Property."
5;"PID_KEYWORDS";"Keywords Summary Property."
6;"PID_COMMENTS";"Comments Summary Property."
7;"PID_TEMPLATE";"Template Summary Property."
8;"PID_LASTAUTHOR";"Last Saved By Summary Property."
9;"PID_REVNUMBER";"Revision Number Summary Property."
10;"PID_EDITTIME";"Not applicable. Reserved."
11;"PID_LASTPRINTED";"Last Printed Summary Property."
12;"PID_CREATE_DTM";"Create Time/Date Summary Property."
13;"PID_LASTSAVE_DTM";"Last Saved Time/Date Summary Property."
14;"PID_PAGECOUNT";"Page Count Summary Property."
15;"PID_WORDCOUNT";"Word Count Summary Property."
16;"PID_CHARCOUNT";"Character Count Summary Property."
17;"PID_THUMBNAIL";"Not applicable. Reserved."
18;"PID_APPNAME";"Creating Application Summary Property."
"@

$MSIInfoSplit = $MSIInfo.Split("`n")

$values = @{}
ForEach ($line in $MSIInfoSplit){
    Write-Warning "Processing $line"
    $found = $line -match '='
    if ($found) {
        $fragments = ($line.Split("="))
        $value = ($fragments[1]).Substring(1)

        $null = $line -replace ' ','' -match '\[\d{1,2}]'
        $key = $matches[0] -replace '[\[\]]',''
        
        $values[$key] = $value
    }
    
    switch ($line) {
        "" { Write-Warning "Skipping empty line..." }
        Default {}
    }
    
}
$values | fl
