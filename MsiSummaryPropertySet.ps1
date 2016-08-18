class MsiSummaryPropertySet {
    [System.Collections.ArrayList] $MsiSummaryProperties
    [string] $MsiInfoExePath = "C:\Program Files (x86)\Windows Kits\10\bin\x86\MsiInfo.exe"

    MsiSummaryPropertySet(){
        $this.setDefaultMsiSummaryProperties()
        if (-not (Test-Path $this.MsiInfoExePath)) {
            Write-Warning "MsiInfo.exe not found. Certain functionality may not be available."
        }
    }

    MsiSummaryPropertySet($MsiInfoExePath){
        $this.setDefaultMsiSummaryProperties()
        $this.MsiInfoExe = $MsiInfoExePath
        if (-not (Test-Path $this.MsiInfoExePath)) {
            Write-Error "Provided MsiInfo.exe not found. Certain functionality may not be available."
        }
    }
    
    setDefaultMsiSummaryProperties() {
        $this.MsiSummaryProperties = @{}
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(0,"PID_DICTIONARY","Not applicable. Reserved."))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(1,"PID_CODEPAGE","Codepage"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(2,"PID_TITLE","Title"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(3,"PID_SUBJECT","Subject"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(4,"PID_AUTHOR","Author"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(5,"PID_KEYWORDS","Keywords"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(6,"PID_COMMENTS","Comments"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(7,"PID_TEMPLATE","Template"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(8,"PID_LASTAUTHOR","Last Saved By"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(9,"PID_REVNUMBER","Revision Number"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(10,"PID_EDITTIME","Not applicable. Reserved."))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(11,"PID_LASTPRINTED","Last Printed"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(12,"PID_CREATE_DTM","Create Time/Date"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(13,"PID_LASTSAVE_DTM","Last Save Time/Date"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(14,"PID_PAGECOUNT","Page Count"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(15,"PID_WORDCOUNT","Word Count"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(16,"PID_CHARCOUNT","Character Count"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(17,"PID_THUMBNAIL","Not applicable. Reserved."))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(18,"PID_APPNAME","Creating Application"))
        $this.MsiSummaryProperties.Add([MsiSummaryProperty]::new(19,"PID_SECURITY","Security"))
    }

    setValueByMsiPID($MsiPID, $Value) {
        ($this.MsiSummaryProperties | ? {$_.MsiPID -eq $MsiPID}).Value = $Value
    }
    
    setValueByPropertyID($PropertyID, $Value) {
        ($this.MsiSummaryProperties | ? {$_.PropertyID -eq $PropertyID}).Value = $Value
    }
    
    loadValuesFromMsi($MsiPath){
        if (-not (Test-Path $this.MsiInfoExePath)) {
            Throw "MsiInfo.exe not found. Cannot load Summary Properties!"
        }
        $MsiInfo = & $this.MsiInfoExePath $MsiPath | Out-String
        $this.setValuesByMsiInfoOutput($MsiInfo)
    }
    
    setValuesByMsiInfoOutput($MSIInfoOutput){
        $MSIInfoSplit = $MSIInfoOutput.Split("`n")
        $MSIInfoSplit | % {$this.setValueByMsiInfoString($_)}
    }
    
    setValueByMsiInfoString($MsiInfoString){
        $found = $MsiInfoString -match '='
        if ($found) {
            $fragments = ($MsiInfoString.Split("="))
            $value = ($fragments[1]).Substring(1)

            $null = $MsiInfoString -replace ' ','' -match '\[\d{1,2}]'
            $key = $matches[0] -replace '[\[\]]',''

            $this.setValueByMsiPID($key,$value)
        }
    }

}