<#
.SYNOPSIS
    Connects with an authenticated account to use Microsoft Graph cmdlet requests.
.DESCRIPTION
    The cmdlet connects an authenticated account to use for Microsoft Graph cmdlet requests.
.EXAMPLE
    PS C:\>Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/User.Read','https://graph.microsoft.com/Files.ReadWrite'
    Authenticate to Microsoft Graph (with permissions User.Read and Files.ReadWrite) using client id from application registration (public client).
.EXAMPLE
    PS C:\>Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -Interactive -Scope 'https://graph.microsoft.com/User.Read' -LoginHint user@domain.com
    Force interactive authentication to Microsoft Graph (with permissions User.Read) for specific Azure AD tenant using client id from application registration (public client).
.EXAMPLE
    PS C:\>Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000' -ClientSecret (ConvertTo-SecureString 'SuperSecretString' -AsPlainText -Force) -Scope 'https://graph.microsoft.com/.default'
    Authenticate to Microsoft Graph (with permissions .Default) using client id and secret from application registration (confidential client).
.EXAMPLE
    PS C:\>$ClientCertificate = Get-Item Cert:\CurrentUser\My\0000000000000000000000000000000000000000
    PS C:\>Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000' -ClientCertificate $ClientCertificate -TenantId '00000000-0000-0000-0000-000000000000'
    Authenticate to Microsoft Graph (with permissions .Default) for specific Azure AD tenant using client id and certificate.
#>
function Connect-MsGraph {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Graph.GraphServiceClient])]
    param
    (
        # Identifier of the client requesting the token.
        [parameter(Mandatory=$true, ParameterSetName='PublicClient')]
        [parameter(Mandatory=$false, ParameterSetName='PublicClient-InputObject')]
        [parameter(Mandatory=$true, ParameterSetName='ConfidentialClientSecret')]
        [parameter(Mandatory=$true, ParameterSetName='ConfidentialClientCertificate')]
        [parameter(Mandatory=$false, ParameterSetName='ConfidentialClient-InputObject')]
        [string] $ClientId,
        # Secure secret of the client requesting the token.
        [parameter(Mandatory=$true, ParameterSetName='ConfidentialClientSecret')]
        [parameter(Mandatory=$false, ParameterSetName='ConfidentialClient-InputObject')]
        [securestring] $ClientSecret,
        # Client assertion certificate of the client requesting the token.
        [parameter(Mandatory=$true, ParameterSetName='ConfidentialClientCertificate')]
        [parameter(Mandatory=$false, ParameterSetName='ConfidentialClient-InputObject')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCertificate,
        # Address to return to upon receiving a response from the authority.
        [parameter(Mandatory=$false)]
        [uri] $RedirectUri,
        # Tenant identifier of the authority to issue token.
        [parameter(Mandatory=$false)]
        [string] $TenantId,
        # Public client application options
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='PublicClient-InputObject', Position=0)]
        [Microsoft.Identity.Client.PublicClientApplicationOptions] $PublicClientOptions,
        # Confidential client application options
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='ConfidentialClient-InputObject', Position=0)]
        [Microsoft.Identity.Client.ConfidentialClientApplicationOptions] $ConfidentialClientOptions,
        # Array of scopes requested for resource
        [Parameter(Mandatory=$false)]
        [string[]] $Scopes = 'https://graph.microsoft.com/.default',
        # Microsoft Graph version
        [Parameter(Mandatory=$false)]
        [ValidateSet('v1.0','beta','canary')]
        [string] $Version
    )

    [hashtable] $paramMsalClientApplication = $PSBoundParameters
    if ($paramMsalClientApplication.ContainsKey('Scopes')) { [void] $paramMsalClientApplication.Remove('Scopes') }
    if ($paramMsalClientApplication.ContainsKey('Version')) { [void] $paramMsalClientApplication.Remove('Version') }
    $script:MsalClientApplication = New-MsalClientApplication @paramMsalClientApplication
    $script:MsalClientApplication | Get-MsalToken -Scopes $Scopes -ErrorAction Stop | Out-Null
    $script:Scopes = $Scopes

    $script:GraphServiceClient = New-MsGraphClient -Version $Version
    return $script:GraphServiceClient
}
