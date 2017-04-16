Function Get-PublicDnsRecord{
    <#
    .SYNOPSIS
        Make some DNS query based on Stat DNS.
    .DESCRIPTION
        Use Invoke-WebRequest on Stat DNS to resolve DNS query.
    .EXAMPLE
        Get-PublicDnsRecord -DomaineNAme "ItForDummies.net" -DnsRecordType A,MX
    .EXAMPLE
        Get-PublicDnsRecord -DomaineNAme "www.valbox.fr" -DnsRecordType A,MX
    .PARAMETER DomainName
        Domain name to query.
    .PARAMETER DnsRecordType
        DNS type to query.
    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
    #>
    Param(
        [Parameter(Mandatory=$true,Position=1)]
        [String]$DomainName,

        [Parameter(Mandatory=$true,Position=2)]
        [ValidateSet('A','AAAA','CERT','CNAME','DHCIP','DLV','DNAME','DNSKEY','DS','HINFO','HIP','IPSECKEY','KX','LOC','MX','NAPTR','NS','NSEC','NSEC3','NSEC3PARAM','OPT','PTR','RRSIG','SOA','SPF','SRV','SSHFP','TA','TALINK','TLSA','TXT')]
        [String[]]$DnsRecordType
    )
    Begin{}
    Process{
        ForEach($Record in $DnsRecordType){
            Try{
                $WebUrl = 'http://www.dns-lg.com/opendns1/{0}/{1}' -f $DomainName,$Record
                
                $WebData = Invoke-WebRequest $WebUrl -ErrorAction Stop | Select-Object -ExpandProperty Content | ConvertFrom-Json | Select-Object -ExpandProperty answer
                $WebData | % {
                     New-Object -TypeName PSObject -Property @{
                        'Name'      = $_.name
                        'Type'      = $_.type
                        'Target'    = $_.rdata
                    }
                }
            }
            catch{
                Write-Warning -Message $_
                New-Object -TypeName PSObject -Property @{
                    'Name'      = $DomainName
                    'Type'      = $Record
                    'Target'    = ($_[0].ErrorDetails.Message -split '"')[-2]
                }
            }
        }
    }
    End{}
}