Function New-GraphToken {
    #Requires -Module AzureRM
    [CmdletBinding()]
    Param(
        $TenantName = 'ItForDummies.net'
    )

    try{
        Import-Module AzureRM -ErrorAction Stop
    }
    catch{
        Write-Error 'Can''t load AzureRM module.'
        break
    }

    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2" #PowerShell ClientID
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.windows.net"
    $authority = "https://login.windows.net/$TenantName"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    #$authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$redirectUri, "Auto")
    $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$redirectUri, "Always")

    @{
       'Content-Type'='application\json'
       'Authorization'=$authResult.CreateAuthorizationHeader()
    }
}