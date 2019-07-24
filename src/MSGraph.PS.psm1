Set-StrictMode -Version 2.0

$MsalClientApplication = $null

[Microsoft.Graph.AuthenticateRequestAsyncDelegate] $AuthenticateRequestAsyncDelegate = {
    param([System.Net.Http.HttpRequestMessage] $requestMessage)
    $MsalToken = $script:MsalClientApplication | Get-MsalToken -Scopes $script:Scopes -ErrorAction Stop
    $requestMessage.Headers.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue ("bearer", $MsalToken.AccessToken)
    return [System.Threading.Tasks.Task]::FromResult(0)
}

$GraphServiceClient = New-Object Microsoft.Graph.GraphServiceClient -ArgumentList (New-Object Microsoft.Graph.DelegateAuthenticationProvider -ArgumentList $AuthenticateRequestAsyncDelegate)
