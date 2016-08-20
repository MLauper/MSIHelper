function Initialize-MsiHelper { 
 
  [CmdletBinding(SupportsShouldProcess=$False)] 
  param 
  ( 
  ) 

  $Script:myMsiSummaryPropertySet = [MsiSummaryPropertySet]::new()
  $Script:myMsiDBPropertySetRequired = [MsiDBPropertySetRequired]::new()
  
}