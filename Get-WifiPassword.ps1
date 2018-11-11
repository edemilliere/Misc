Function Get-WifiPassword {
    netsh wlan show profile | Select-Object -Skip 3| Where-Object -FilterScript {($_ -like '*:*')} | ForEach-Object -Process {
        $NetworkName = $_.Split(':')[-1].trim()
        $PasswordDetection = $(netsh wlan show profile name =$NetworkName key=clear) | Where-Object -FilterScript {($_ -like '*contenu de la clé*') -or ($_ -like '*key content*')}

        New-Object -TypeName PSObject -Property @{
            NetworkName = $NetworkName
            Password = if($PasswordDetection){$PasswordDetection.Split(':')[-1].Trim()}else{'Unknown'}
        } -ErrorAction SilentlyContinue
    } 
}