<#
.SYNOPSIS
    Invoke Microsoft Graph request.
.DESCRIPTION

.EXAMPLE
    PS C:\>$MsGraphClient = Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000' -Scope 'https://graph.microsoft.com/User.Read'
    PS C:\>$MsGraphClient.Me | Invoke-MsGraphRequest
    Connects Microsoft Graph client with User.Read scope and invokes request on Me endpoint.
.EXAMPLE
    PS C:\>$MsGraphClient = Connect-MsGraph -ClientId '00000000-0000-0000-0000-000000000000'
    PS C:\>$MsGraphClient.Users.Request().Filter("DisplayName eq 'Joe Cool'") | Invoke-MsGraphRequest -Scope 'https://graph.microsoft.com/User.Read.All'
    Connects Microsoft Graph client and invokes request on User endpoint with filter and specific scope.
#>
function Invoke-MsGraphRequest {
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    param
    (
        # Client application options
        [parameter(Mandatory=$false, ValueFromPipeline=$true, ParameterSetName='InputObject', Position=0)]
        [object] $InputObject,
        # Array of scopes requested for resource
        [Parameter(Mandatory = $false)]
        [string[]] $Scopes
    )

    ## Check for authentication
    if (!$script:MsalClientApplication) {
        # Write a terminating error message indicating the user must authenticate.
        $errorMessage = "You must call the Connect-MsGraph cmdlet before calling any other cmdlets."
        Write-Error -Message $errorMessage -Category ([System.Management.Automation.ErrorCategory]::AuthenticationError) -ErrorId "InvokeMsGraphRequestFailureNeedAuthentication" -ErrorAction Stop
    }

    ## Update Global Scopes
    if ($Scopes) { $script:Scopes = $Scopes }

    ## Invoke Request
    switch ($PSCmdlet.ParameterSetName) {
        "InputObject" {
            ## InputObject Casting
            if ($InputObject -is [Microsoft.Graph.BaseRequestBuilder]) {
                $Result = $InputObject.Request().GetAsync().GetAwaiter().GetResult()
            }
            elseif ($InputObject -is [Microsoft.Graph.BaseRequest]) {
                $Result = $InputObject.GetAsync().GetAwaiter().GetResult()
            }
            elseif ($InputObject -is [System.Net.Http.HttpRequestMessage]) {
                $script:GraphServiceClient.AuthenticationProvider.AuthenticateRequestAsync($InputObject)
                $Result = $script:GraphServiceClient.HttpProvider.SendAsync($InputObject)
            }
            else {
                # Otherwise, write a terminating error message indicating that input object type is not supported.
                $errorMessage = "Cannot invoke request on type {0}." -f $InputObject.GetType()
                Write-Error -Message $errorMessage -Category ([System.Management.Automation.ErrorCategory]::ParserError) -ErrorId "InvokeMsGraphRequestFailureTypeNotSupported" -ErrorAction Stop
            }
        }
    }

    return $Result
}
