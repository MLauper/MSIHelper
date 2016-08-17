$winst = New-Object -ComObject "WindowsInstaller.Installer"

$winst | gm

$winst.InitializeLifetimeService() | gm

$winsttype = $winst.GetType()

$winsttype | gm

$winsttype.InvokeMember("Programs", 
    [System.Reflection.BindingFlags]::InvokeMethod,
    $null, $winst, $null, $null, $null, $null, $null)

$winst.GetType().InvokeMember(
    "OpenDatabase", 
    [System.Reflection.BindingFlags]::InvokeMethod,
    $null,  ## Binder
    $winst,  ## Target
    $null,  ## Args
    $null,  ## Modifiers
    $null,  ## Culture
    $null  ## NamedParameters
            )

$type.invokeMember($args[0],[System.Reflection.BindingFlags]::InvokeMethod,$null,$this,$methodargs)

$shell = New-Object -ComObject Shell.Application
Invoke-NamedParameter $Shell "Explore" @{"vDir"="$pwd"}

$winst = New-Object -ComObject WindowsInstaller.Installer -Strict
Invoke-NamedParameter $winst "OpenDatabase" @{"path"="C:\Users\Marco\Desktop\138a31.msi";"whatever"="0"}

Function Invoke-NamedParameter {
    [CmdletBinding(DefaultParameterSetName = "Named")]
    param(
        [Parameter(ParameterSetName = "Named", Position = 0, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 0, Mandatory = $true)]
        [ValidateNotNull()]
        [System.Object]$Object
        ,
        [Parameter(ParameterSetName = "Named", Position = 1, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Method
        ,
        [Parameter(ParameterSetName = "Named", Position = 2, Mandatory = $true)]
        [ValidateNotNull()]
        [Hashtable]$Parameter
        ,
        [Parameter(ParameterSetName = "Positional")]
        [Object[]]$Argument
    )

    end {  ## Just being explicit that this does not support pipelines
        if ($PSCmdlet.ParameterSetName -eq "Named") {
            ## Invoke method with parameter names
            ## Note: It is ok to use a hashtable here because the keys (parameter names) and values (args)
            ## will be output in the same order.  We don't need to worry about the order so long as
            ## all parameters have names
            $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod,
                $null,  ## Binder
                $Object,  ## Target
                ([Object[]]($Parameter.Values)),  ## Args
                $null,  ## Modifiers
                $null,  ## Culture
                ([String[]]($Parameter.Keys))  ## NamedParameters
            )
        } else {
            ## Invoke method without parameter names
            $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod,
                $null,  ## Binder
                $Object,  ## Target
                $Argument,  ## Args
                $null,  ## Modifiers
                $null,  ## Culture
                $null  ## NamedParameters
            )
        }
    }
}









# $vw.Execute()
$vw = $vw | Add-Member -MemberType ScriptMethod -Value $codeInvokeMethod -Name InvokeMethod -PassThru
$vw.InvokeMethod("Execute")
 
# $rec = $vw.Fetch()
$rec = $vw.InvokeMethod("Fetch")
 
$codeInvokeParamProperty = {
    $type = $this.gettype();
    $index = $args.count -1 ;
    $methodargs=$args[1..$index]
    $type.invokeMember($args[0],[System.Reflection.BindingFlags]::GetProperty,$null,$this,$methodargs)
}
 
$rec = $rec | Add-Member -MemberType ScriptMethod -Value $codeInvokeParamProperty -Name InvokeParamProperty -PassThru
 
If ($rec -ne $null)
{
     
    $DataSize = $rec.InvokeParamProperty("DataSize",2)
 
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa371140%28v=vs.85%29.aspx
    $paramHT = @{Field = 2 ; Length = [int]$DataSize ; Format = 1}
 
    $sMetadata = $rec.GetType().InvokeMember("ReadStream", [System.Reflection.BindingFlags]::InvokeMethod,
                $null,  # Binder
                $rec,  # Target
                ([Object[]]($paramHT.Values)),  # Args
                $null,  # Modifiers
                $null,  # Culture
                ([String[]]($paramHT.Keys))  # NamedParameters
            )
     
} else {
    Write-Output -InputObject "No Metadata stream was found in this file: $Path"
    Exit 2
}