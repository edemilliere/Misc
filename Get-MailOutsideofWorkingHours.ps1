<#
    .SYNOPSIS
        This script will use message tracking logs to search for mails sent or received outside of working hours.
    .DESCRIPTION
        Working hours are 8 to 8 on week days only.
    .EXAMPLE
        .\Get-MailOutsideofWorkingHours.ps1 -UserList @('dumbo@itfordummies.net','junk@itfordummies.net') | Out-GridView -Title 'Mails'
    .EXAMPLE
        .\Get-MailOutsideofWorkingHours.ps1 -UserList @('dumbo@itfordummies.net','junk@itfordummies.net') | Where-Object -FilterScript {$_.OutsideOfWorkingHours -eq $true} | Out-GridView -Title 'Mail Outside of Working Hours !'
    .PARAMETER UserList
        List of mailboxes to search.
    .PARAMETER StartDate
        Start date for the search.

        Defaulted to 7 days ago.
    .PARAMETER EndDate
        End date for the search.

        Defaulted to now.
    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
        http://ItForDummies.net
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'List of send/Recipient to check.',
        Position = 0)]
    [String[]]$UserList,

    [Parameter(Mandatory = $false,
        HelpMessage = 'Start date for the mail search.',
        Position = 1)]
    [DateTime]$StartDate = (Get-Date).AddDays(-7),
    
    [Parameter(Mandatory = $false,
        HelpMessage = 'End date for the mail search.',
        Position = 2)]
    [DateTime]$EndDate = (Get-Date),

    [PSCredential]$Credential = (Get-Credential)
)

#region Helpers
Function Connect-ExchangeOnline{
    [CmdletBinding()]
    Param(
        [PSCredential]$Credential = (Get-Credential)
    )
    Import-Module (Import-PSSession -Session $(New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credential -Authentication Basic -AllowRedirection) -DisableNameChecking) -Global -DisableNameChecking
}
#endregion

Connect-ExchangeOnline -Credential $Credential

#region Main
ForEach($User in $UserList){
    #region Recipient
    $Continue = $true
    $CurrentPage = 1
    do{
        $CurrentMessageTrace = Get-MessageTrace -RecipientAddress $User -StartDate $StartDate -EndDate $EndDate -PageSize 100 -Page $CurrentPage
        if($CurrentMessageTrace){
            $CurrentMessageTrace | Select-Object -Property 'SenderAddress','RecipientAddress','Subject','Received',@{Label='OutsideOfWorkingHours';Expression={if(($_.Received.DayOfWeek -eq [System.DayOfWeek]::Saturday) -or ($_.Received.DayOfWeek -eq [System.DayOfWeek]::Sunday) -or ($_.Received.TimeOfDay.Hours -lt 8) -or ($_.Received.TimeOfDay.Hours -gt 18) ){$true}else{$false}}}
        }
        else{$Continue = $false}
        $CurrentMessageTrace = $null

        $CurrentPage++
    }
    while($Continue)
    #endregion

    #region Sender
    $Continue = $true
    $CurrentPage = 1
    do{
        $CurrentMessageTrace = Get-MessageTrace -SenderAddress $User -StartDate $StartDate -EndDate $EndDate -PageSize 100 -Page $CurrentPage
        if($CurrentMessageTrace){
            $CurrentMessageTrace | Select-Object -Property 'SenderAddress','RecipientAddress','Subject','Received',@{Label='OutsideOfWorkingHours';Expression={if(($_.Received.DayOfWeek -eq [System.DayOfWeek]::Saturday) -or ($_.Received.DayOfWeek -eq [System.DayOfWeek]::Sunday) -or ($_.Received.TimeOfDay.Hours -lt 8) -or ($_.Received.TimeOfDay.Hours -gt 18) ){$true}else{$false}}}
        }
        else{$Continue = $false}
        $CurrentMessageTrace = $null

        $CurrentPage++
    }
    while($Continue)
    #endregion
}
#endregion