param
(
    #
    [parameter(Mandatory=$false)]
    [string] $ModuleManifestPath = ".\src\MSGraph.PS.psd1"
)

#.\build\Restore-NugetPackages.ps1 -BaseDirectory ".\" -Verbose:$false
Import-Module $ModuleManifestPath
