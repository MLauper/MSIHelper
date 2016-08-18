class MsiSummaryProperty {
    [int16] $MsiPID
    [string] $PropertyID
    [string] $PropertyName
    [string] $Value
    MsiSummaryProperty([int16] $MsiPID, [string] $PropertyID, [string] $PropertyName) {
        $this.MsiPID = $MsiPID
        $this.PropertyID = $PropertyID
        $this.PropertyName = $PropertyName
    }
    MsiSummaryProperty([int16] $MsiPID, [string] $PropertyID, [string] $PropertyName, [string] $Value) {
        $this.MsiPID = $MsiPID
        $this.PropertyID = $PropertyID
        $this.PropertyName = $PropertyName
        $this.Value = $Value
    }
}