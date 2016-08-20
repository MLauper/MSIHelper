class MsiDBPropertyEntry {
    [string] $Property
    [string] $Value
    MsiDBPropertyEntry([string] $Property) {
        $this.Property = $Property
    }
    MsiDBPropertyEntry([string] $Property, [string] $Value) {
        $this.Property = $Property
        $this.Value = $Value
    }
}