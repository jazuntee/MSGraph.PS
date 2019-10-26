param
(
	#
    [parameter(Mandatory=$false)]
    [string] $ModulePath = ".\release\MSGraph.PS\1.18.0.1",
    #
    [parameter(Mandatory=$true)]
    [string] $NuGetApiKey
)

Publish-Module -Path $ModulePath -NuGetApiKey $NuGetApiKey
