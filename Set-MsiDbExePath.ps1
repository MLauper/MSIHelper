function Set-MsiDbExePath { 
  <# 
  .SYNOPSIS 
  Set MsiDb.exe location to be used for all subsequent calls.  
  .DESCRIPTION 
  This function sets the location of the MsiDb.exe. All subsequent calls will use this exe, until another path is specified.
  The setting is valid for the duration of the current session.
  .EXAMPLE 
  Set-MsiInfoExePath "C:\TEMP\MsiDb.exe"  
  .PARAMETER MsiInfoExePath
  The MsiDb.exe executable which should be used in subsequent calls.
  #> 
  [CmdletBinding(SupportsShouldProcess=$False)] 
  param 
  ( 
    [Parameter(Mandatory=$False, Position=2)]
    [ValidateScript({Test-Path $_ -PathType "Leaf"})]
    [string]$MsiDbExePath
  ) 

  $Script:myMsiDBPropertySetRequired.setMsiDbExePath($MsiDbExePath)

  return
}