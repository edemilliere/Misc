#requires -modules ExchangeOnlineManagement

#Can be one or the other
$Mailboxes = 'junk@itfordummies.net','dumbo@itfordummies.net'
#$Mailboxes = Get-Content MailboxesList.txt

#Exchange Online Connection, can beinstalled with:
#Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline

#Main loop
$Output = foreach($Mailbox in $Mailboxes){
    
    $OldestItem = Get-MailboxFolderStatistics -Identity $Mailbox -IncludeOldestAndNewestItems | Select-Object -ExpandProperty OldestItemReceivedDate | Sort-Object | Select-Object -First 1
    $TotalSizeInMb = ("$(Get-MailboxStatistics -Identity $Mailbox | Select-Object -ExpandProperty TotalItemSize | Select-Object -ExpandProperty Value)" -split '\(| ')[-2].Replace(',','')/1MB -as [int]
    $AgeInDays = (Get-Date) - $OldestItem | Select-Object -ExpandProperty Days
    $DailyaverageInMB = $TotalSizeInMb/$AgeInDays
    
    [PSCustomObject]@{
        Identity = $Mailbox
        OldestItem = $OldestItem
        TotalSizeInMB = $TotalSizeInMb
        AgeInDays = $AgeInDays
        DailyaverageInMB = '{0:N2}' -f $DailyaverageInMB
        '30Daysaverage' = ($DailyaverageInMB*30) -as [int]
    }
    
    $OldestItem = $TotalSizeInMb = $AgeInDays = $DailyaverageInMB = $null
}

$Output | Export-Csv -NoTypeInformation -Delimiter ';' 30DaysOfEmailCalculatedFromTotalSizeAndLifeTimeAverage.csv