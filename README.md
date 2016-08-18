# MSIHelper

This is a simple PowerShell Module that retrieves the MSI Summary Property Set.

All retrieved attributes are documented here: https://msdn.microsoft.com/en-us/library/windows/desktop/aa372045(v=vs.85).aspx

The Module uses the tool MsiInfo.exe and parses the command line. This has been done, because Microsoft seems to provide only COM-Objects, which seem to be a bit buggy, and C++ bindings, which I didn't want to use.

### Prerequisites
The module requires MsiInfo.exe (https://msdn.microsoft.com/en-us/library/windows/desktop/aa370310(v=vs.85).aspx) and PowerShell 5.0.

MSIInfo is available as part of the Windows SDK which is available for free from https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk. 
PowerShell 5.0 is available as part the WMF 5 which is as well available for free from https://www.microsoft.com/en-us/download/details.aspx?id=50395.

### Usage
You can either copy the module in a standard module directory or import it directly from the cloned repo. 
If you have installed the Windows 10 SDK in the default location, you can run: 

```
git clone https://github.com/MLauper/MSIHelper.git
Import-Module .\MSIHelper\MSIHelper.psd1
Get-MsiInfo "C:\Program Files (x86)\Windows Kits\10\bin\x86\Orca-x86_en-us.msi"
```
Will result in:
```
MsiPID PropertyID       PropertyName              Value
------ ----------       ------------              -----
     0 PID_DICTIONARY   Not applicable. Reserved.
     1 PID_CODEPAGE     Codepage                  1252...
     2 PID_TITLE        Title                     Installation Database...
     3 PID_SUBJECT      Subject                   Orca...
     4 PID_AUTHOR       Author                    Microsoft Corporation...
     5 PID_KEYWORDS     Keywords                  Installer...
     6 PID_COMMENTS     Comments                  This installer database contains the logic and data required to in...
     7 PID_TEMPLATE     Template                  Intel;1033...
     8 PID_LASTAUTHOR   Last Saved By
     9 PID_REVNUMBER    Revision Number           {6EBA87D8-9142-4C2A-902D-6F3B86340D44}...
    10 PID_EDITTIME     Not applicable. Reserved.
    11 PID_LASTPRINTED  Last Printed
    12 PID_CREATE_DTM   Create Time/Date          2016/07/29 06:18:14...
    13 PID_LASTSAVE_DTM Last Save Time/Date       2016/07/29 06:18:14...
    14 PID_PAGECOUNT    Page Count                400...
    15 PID_WORDCOUNT    Word Count                2...
    16 PID_CHARCOUNT    Character Count
    17 PID_THUMBNAIL    Not applicable. Reserved.
    18 PID_APPNAME      Creating Application      Windows Installer XML (3.7.4128.0)...
    19 PID_SECURITY     Security                  2...
```