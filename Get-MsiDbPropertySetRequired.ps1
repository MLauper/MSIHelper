function Get-MsiDbPropertySetRequired { 
  <# 
  .SYNOPSIS 
  Get required values from the Property table from the Msi Database.
  .DESCRIPTION 
  This function retrieves all required values ffrom the Property table from the Msi Database.
  These values are documented under https://msdn.microsoft.com/en-us/library/windows/desktop/aa370905(v=vs.85).aspx.
  The retrieved properties are ProductLanguage, Manufacturer, ProductCode, ProductName and ProductVersion
  .EXAMPLE 
  Get-MsiDbPropertySetRequired "C:\Program Files (x86)\Windows Kits\10\bin\x86\Orca-x86_en-us.msi"
  .PARAMETER MsiFilePath 
  The Msi from which the information is retrieved.  
  #> 
  [CmdletBinding(SupportsShouldProcess=$False)] 
  param 
  ( 
    [Parameter(Mandatory=$True, Position=1)]
    [ValidateScript({Test-Path $_ -PathType "Leaf"})]
    [string]$MsiFilePath
  )

  $Script:myMsiDBPropertySetRequired.loadValuesFromMsi($MsiFilePath)
  
  return $Script:myMsiDBPropertySetRequired.MsiDBPropertySetRequired

}