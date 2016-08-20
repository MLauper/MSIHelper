function Get-MsiInfo { 
  <# 
  .SYNOPSIS 
  Get Summary Info from MSI. 
  .DESCRIPTION 
  This function retrieves the  summary informations from a MSI package.
  All Values are described here: https://msdn.microsoft.com/en-us/library/windows/desktop/aa372045(v=vs.85).aspx
  .EXAMPLE 
  Get-MsiInfo "C:\Program Files (x86)\Windows Kits\10\bin\x86\Orca-x86_en-us.msi"
  .EXAMPLE 
  Get-MsiInfo -MsiFilePath "C:\Program Files (x86)\Windows Kits\10\bin\x86\Orca-x86_en-us.msi"
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

  $Script:myMsiSummaryPropertySet.loadValuesFromMsi($MsiFilePath)
  
  return $Script:myMsiSummaryPropertySet.MsiSummaryProperties

}