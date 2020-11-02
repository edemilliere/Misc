﻿#requires -modules ExchangeOnlineManagement
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $true)
    ]
    $MailboxesToSearchIn = ('junk@itfordummies.net','toto@itfordummies.net')
)

#region Prereqs
#Exchange Online Connection, can beinstalled with:
#Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Connect-IPPSSession
#endregion

#region Init
[Regex]$Regex = "Location:\s(?<Identity>\w+@\w+.\w+),\sItem\scount:\s(?<ItemCount>\w+),\sTotal\ssize:\s(?<TotalSize>\w+)"

$ComplianceSearchName = "30Days - $(Get-Date)"

$StartDate = (Get-Date).AddDays(-30).ToString("MM-dd-yyyy")
$EndDate = (Get-Date).ToString("MM-dd-yyyy")
$ContentMatchQuery = "(sent=$StartDate..$EndDate)(received=$StartDate..$EndDate)"
#endregion

#Create compliance search
New-ComplianceSearch -ExchangeLocation $MailboxesToSearchIn -ContentMatchQuery $ContentMatchQuery -Name $ComplianceSearchName

#Trigger search
Get-ComplianceSearch -Identity $ComplianceSearchName | Start-ComplianceSearch

#Wait a few minutes to let it complete
do{
    Write-Host -NoNewline '.'
    Start-Sleep -Seconds 10
}
while((Get-ComplianceSearch -Identity $ComplianceSearchName).status -ne 'Completed')

#Get the result
$SuccessResults = (Get-ComplianceSearch -Identity $ComplianceSearchName | Select-Object -ExpandProperty SuccessResults) #-replace '\r\n'
$RegexMatchingResult = $Regex.Matches($SuccessResults)
$RegexMatchingResult | ForEach-Object -Process {
    [PSCustomObject]@{
        Identity = $_.Groups['Identity'].Value
        ItemCount = $_.Groups['ItemCount'].Value        
        TotalSizeinMB = $_.Groups['TotalSize'].Value/1MB -as [int]
    }
} | Export-Csv -NoTypeInformation -Delimiter ';' "30DaysOfEmailCalculatedFromComplianceSearch-$(Get-Date).csv"

#Cleanup
Get-ComplianceSearch -Identity $ComplianceSearchName | Remove-ComplianceSearch -Confirm:$false