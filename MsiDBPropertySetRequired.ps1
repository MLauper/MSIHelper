class MsiDBPropertySetRequired {
    [System.Collections.ArrayList] $MsiDBPropertySetRequired
    [string] $MsiDbExePath = "C:\Program Files (x86)\Windows Kits\10\bin\x86\MsiDb.exe"

    MsiDBPropertySetRequired(){
        $this.initializeRequiredMsiDbProperties()
    }

    MsiDBPropertySetRequired($MsiDbExePath){
        $this.initializeRequiredMsiDbProperties()
        $this.setMsiDbExePath($MsiDbExePath)
    }
    
    initializeRequiredMsiDbProperties() {
        $this.MsiDBPropertySetRequired = @{}
        $this.MsiDBPropertySetRequired.Add([MsiDBPropertyEntry]::new("ProductLanguage"))
        $this.MsiDBPropertySetRequired.Add([MsiDBPropertyEntry]::new("Manufacturer"))
        $this.MsiDBPropertySetRequired.Add([MsiDBPropertyEntry]::new("ProductCode"))
        $this.MsiDBPropertySetRequired.Add([MsiDBPropertyEntry]::new("ProductName"))
        $this.MsiDBPropertySetRequired.Add([MsiDBPropertyEntry]::new("ProductVersion"))
    }

    setValueByPropertyIdentifier($PropertyIdentifier, $Value) {
        ($this.MsiDBPropertySetRequired | ? {$_.Property -eq $PropertyIdentifier}).Value = $Value
    }
    
    loadValuesFromMsi($MsiPath){
        $this.initializeRequiredMsiDbProperties()
        if (-not (Test-Path $this.MsiDbExePath)) {
            Throw "MsiDb.exe not found. Cannot load database entries! `nCurrent path: " + $this.MsiDbExePath + "`nUse Set-MsiDbExePath `$MsiDbExePath to set the proper location of the module."
        }

        $DumpPath = $this.dumpPropertyTable($MsiPath)

        $PropertyTableDumpPath = Join-Path $DumpPath "Property.idt"

        if (-not (Test-Path -PathType Leaf $PropertyTableDumpPath)) {
            throw "Property table could not be written.`nOutput Path: $PropertyTableDumpPath"
        }

        $PropertyTableDump = Get-Content $PropertyTableDumpPath
        Remove-Item $DumpPath -Recurse -Force 
        
        $this.setValuesByPropertyTableDump($PropertyTableDump)
    }

    [string] dumpPropertyTable($MsiPath){
        $TempPath = [System.IO.Path]::GetTempPath()
        $TempDir = [System.IO.Path]::GetRandomFileName()
        $DumpPath = Join-Path $TempPath $TempDir

        New-Item -ItemType Directory -Path $DumpPath
        
        & $this.MsiDbExePath -e -d $MsiPath -f $DumpPath Property | Out-Null

        return $DumpPath
    }

    setValuesByPropertyTableDump($PropertyTableDump){
        $PropertyTableDumpSplit = $PropertyTableDump.Split("`n")
        $PropertyTableDumpSplit | % {$this.setValueByIdtString($_)}
    }
    
    setValueByIdtString($MsiInfoString){
        if ($MsiInfoString -match '(?<Property>.*)\t(?<Value>.*)') {
            if ($this.MsiDBPropertySetRequired | ? {$_.Property -eq $matches["Property"]}){
                $this.setValueByPropertyIdentifier($matches["Property"], $matches["Value"])
            }
        }
    }

    setMsiDbExePath($MsiDbExePath){
        $this.MsiDbExePath = $MsiDbExePath
        if (-not (Test-Path $this.MsiDbExePath)) {
            Write-Error "Provided MsiDb.exe not found. Certain functionality may not be available. `nProvided path: $MsiDbExePath `nUse Set-MsiDbExePath `$MsiDbExePath to set the proper location of the module."
        }
    }
}