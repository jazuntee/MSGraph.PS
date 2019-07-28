Set-StrictMode -Version 2.0

## PowerShell Desktop 5.1 does not dot-source ScriptsToProcess when a specific version is specified on import. This is a bug.
if ($PSEdition -eq 'Desktop') {
    $ModuleManifest = Import-PowershellDataFile (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.psm1','.psd1'))
    if ($ModuleManifest.ContainsKey('ScriptsToProcess')) {
        foreach ($Path in $ModuleManifest.ScriptsToProcess) {
            . (Join-Path $PSScriptRoot $Path)
        }
    }
}

##
[Microsoft.Identity.Client.IClientApplicationBase] $MsalClientApplication = $null

[Microsoft.Graph.AuthenticateRequestAsyncDelegate] $AuthenticateRequestAsyncDelegate = {
    param([System.Net.Http.HttpRequestMessage] $requestMessage)
    $MsalToken = $script:MsalClientApplication | Get-MsalToken -Scopes $script:Scopes -ErrorAction Stop
    $requestMessage.Headers.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue ("bearer", $MsalToken.AccessToken)
    return [System.Threading.Tasks.Task]::FromResult(0)
}

$GraphServiceClient = New-Object Microsoft.Graph.GraphServiceClient -ArgumentList (New-Object Microsoft.Graph.DelegateAuthenticationProvider -ArgumentList $AuthenticateRequestAsyncDelegate)
