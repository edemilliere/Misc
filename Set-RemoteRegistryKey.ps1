Function Set-RemoteRegistryKey {
    <#
    .SYNOPSIS
        Set registry key on remote computers.
    .DESCRIPTION
        This function uses .Net class [Microsoft.Win32.RegistryKey].
    .PARAMETER ComputerName
        Name of the remote computers.
    .PARAMETER Hive
        Hive where the key is.
    .PARAMETER KeyPath
        Path of the key.
    .PARAMETER Name
        Name of the key setting.
    .PARAMETER Type
        Type of the key setting.
    .PARAMETER Value
        Value tu put in the key setting.
    .EXAMPLE
        Set-RemoteRegistryKey -ComputerName $env:ComputerName -Hive "LocalMachine" -KeyPath "software\corporate\master\Test" -Name "TestName" -Type String -Value "TestValue" -Verbose
    .LINK
        http://itfordummies.net
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage='Provide a ComputerName')]
        [String[]]$ComputerName=$env:ComputerName,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]$Hive,
        
        [Parameter(Mandatory=$true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory=$true)]
        [String]$Name,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryValueKind]$Type,
        
        [Parameter(Mandatory=$true)]
        [Object]$Value
    )
    Begin{
    }
    Process{
        ForEach ($Computer in $ComputerName) {
            try {
                Write-Verbose "Trying computer $Computer"
                $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", "$Computer")
                Write-Debug -Message "Contenur de Reg $reg"
                $key=$reg.OpenSubKey("$KeyPath",$true)
                if($key -eq $null){
                    Write-Verbose -Message "Key not found."
                    Write-Verbose -Message "Calculating parent and child paths..."
                    $parent = Split-Path -Path $KeyPath -Parent
                    $child = Split-Path -Path $KeyPath -Leaf
                    Write-Verbose -Message "Creating the subkey $child in $parent..."
                    $Key=$reg.OpenSubKey("$parent",$true)
                    $Key.CreateSubKey("$child") | Out-Null
                    Write-Verbose -Message "Setting $value in $KeyPath"
                    $key=$reg.OpenSubKey("$KeyPath",$true)
                    $key.SetValue($Name,$Value,$Type)
                }
                else{
                    Write-Verbose "Key found, setting $Value in $KeyPath..."
                    $key.SetValue($Name,$Value,$Type)
                }
                Write-Verbose "$Computer done."
            }#End Try
            catch {Write-Warning "$Computer : $_"} 
        }#End ForEach
    }#End Process
    End{
    }
}