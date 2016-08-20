function Set-MsiInfoExePath { 
  <# 
  .SYNOPSIS 
  Set MsiInfo.exe location to be used for all subsequent calls.  
  .DESCRIPTION 
  This function sets the location of the MsiInfo.exe. All subsequent calls will use this exe, until another path is specified.
  The setting is valid for the duration of the current session.
  .EXAMPLE 
  Set-MsiInfoExePath "C:\TEMP\MsiInfo.exe"  
  .PARAMETER MsiInfoExePath
  The MsiInfo.exe executable which should be used in subsequent calls.
  #> 
  [CmdletBinding(SupportsShouldProcess=$False)] 
  param 
  ( 
    [Parameter(Mandatory=$False, Position=2)]
    [ValidateScript({Test-Path $_ -PathType "Leaf"})]
    [string]$MsiInfoExePath
  ) 

  $Script:myMsiSummaryPropertySet.setMsiInfoExePath($MsiInfoExePath)

  return
}