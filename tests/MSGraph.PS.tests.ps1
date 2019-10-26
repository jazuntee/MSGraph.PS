
Import-Module ..\**\CommonFunctions.psm1

Remove-Module MSGraph.PS -ErrorAction SilentlyContinue
Import-Module ..\src\MSGraph.PS.psd1

### Parameters
[string] $TenantId = 'jasoth.onmicrosoft.com'
[uri] $RedirectUri = 'https://login.microsoftonline.com/common/oauth2/nativeclient'

### Test PublicClient
[string] $PublicClientId = '5f2d068e-50f9-4f92-b532-c62fa531de1f'
[string[]] $DelegatedScopes = @(
    #'https://graph.microsoft.com/.default'
    'https://graph.microsoft.com/User.Read'
)

## Test Public Client Automatic
$MsGraph = Connect-MsGraph -ClientId $PublicClientId -TenantId $TenantId -Scopes $DelegatedScopes -Verbose
$MsGraph.Me | Invoke-MsGraphRequest
$MsGraph.Users.Request().Filter("DisplayName eq 'Jason Thompson'") | Invoke-MsGraphRequest -Scopes 'https://graph.microsoft.com/User.Read.All'


### Test ConfidentialClient
[string] $ConfidentialClientId = 'f3cd10d2-c84d-4b4d-97b5-73109ccef55d'
[securestring] $ConfidentialClientSecret = ConvertTo-SecureString 'SuperSecretString' -AsPlainText -Force
[System.Security.Cryptography.X509Certificates.X509Certificate2] $ConfidentialClientCertificate = Get-Item Cert:\CurrentUser\My\b12afe95f226d94dd01d3f61ae3dbb1c4947ef62
[string[]] $Scopes = @(
    'https://graph.microsoft.com/.default'
    #'https://graph.microsoft.com/User.Read.All'
)

if ($MsalToken.AccessToken) {
    ## Create New Confidential Client?
    [string] $ConfidentialClientId = New-AzureADApplicationConfidentialClient $MsalToken | Select-Object appId
    ## Reset ClientSecret?
    [securestring] $ConfidentialClientSecret = Add-AzureADApplicationClientSecret $MsalToken $ConfidentialClientId
    ## Reset ClientCertificate?
    $ConfidentialClientCertificate = Add-AzureADApplicationClientCertificate $MsalToken $ConfidentialClientId
}

## Test Confidential Client Secret
$MsGraph = Connect-MsGraph -ClientId $ConfidentialClientId -ClientSecret $ConfidentialClientSecret -TenantId $TenantId -Scopes $Scopes -Verbose
## Test Confidential Client Certificate
$MsGraph = Connect-MsGraph -ClientId $ConfidentialClientId -ClientCertificate $ConfidentialClientCertificate -TenantId $TenantId -Scopes $Scopes -Verbose


### Cleanup
## Clear Consent
Get-AzureADServicePrincipal -Filter "AppId eq '$PublicClientId'" | Get-AzureADServicePrincipalOAuth2PermissionGrant | Remove-AzureADOAuth2PermissionGrant

## Remove Certificates from Certificate Store
Get-ChildItem Cert:\CurrentUser\My | Where-Object Subject -eq "CN=ConfidentialClient" | Remove-Item
