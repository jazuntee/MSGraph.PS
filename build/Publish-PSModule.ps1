param
(
	#
    [parameter(Mandatory=$false)]
    [string] $ModulePath = ".\release\MSGraph.PS\1.16.0.3",
    #
    [parameter(Mandatory=$true)]
    [string] $NuGetApiKey
)

Publish-Module -Path $ModulePath -NuGetApiKey $NuGetApiKey
