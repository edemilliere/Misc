#requires -modules ExchangeOnlineManagement
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $false)
    ]
    $Mailboxes
)
#region functions
Function Show-FilePicker{
    Param(
        [String]$InitialDirectory = $pwd,
        [String]$Title = 'Select the CSV file'
    )

    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null

    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = 'All files (*.txt)| *.txt'
    $OpenFileDialog.ShowDialog() | Out-Null
    #return
    $OpenFileDialog.filename
}
#endregion

#input
if(!$Mailboxes){
    $InputFile = Show-FilePicker
    $Mailboxes = Get-Content -Path $InputFile
}

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

$Output | Export-Csv -NoTypeInformation -Delimiter ';' "30DaysOfEmailCalculatedFromTotalSizeAndLifeTimeAverage-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"