<#
.SYNOPSIS
    Creates new Microsoft Graph client object.
.DESCRIPTION

.EXAMPLE
    PS C:\>New-MsGraphClient
    Creates new client with default settings.
.EXAMPLE
    PS C:\>New-MsGraphClient 'beta'
    Creates new client targeting the beta endpoint.
#>
function New-MsGraphClient {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.AuthenticationResult])]
    param
    (
        # Microsoft Graph version
        [Parameter(Mandatory=$false, Position=1)]
        [ValidateSet('','v1.0','beta','canary')]
        [string] $Version
    )

    [uri] $BaseUri = $null
    switch ($Version) {
        'v1.0' { $BaseUri = "https://graph.microsoft.com/$Version" }
        'beta' { $BaseUri = "https://graph.microsoft.com/$Version" }
        'canary' { $BaseUri = "https://$Version.graph.microsoft.com" }
    }

    if ($BaseUri) {
        $GraphServiceClient = New-Object Microsoft.Graph.GraphServiceClient -ArgumentList $BaseUri,(New-Object Microsoft.Graph.DelegateAuthenticationProvider -ArgumentList $script:AuthenticateRequestAsyncDelegate)
    }
    else {
        $GraphServiceClient = New-Object Microsoft.Graph.GraphServiceClient -ArgumentList (New-Object Microsoft.Graph.DelegateAuthenticationProvider -ArgumentList $script:AuthenticateRequestAsyncDelegate)
    }

    return $GraphServiceClient
}
