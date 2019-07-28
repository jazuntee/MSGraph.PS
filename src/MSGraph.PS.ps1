
## Read Module Manifest
$ModuleManifest = Import-PowershellDataFile (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.ps1','.psd1'))
[System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]

## Select the correct assemblies for the PowerShell platform
if ($PSEdition -eq 'Desktop') {
    foreach ($Path in ($ModuleManifest.FileList -like "*\System.ValueTuple.*\portable-net40+sl4+win8+wp8\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Newtonsoft.Json.*\net45\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Microsoft.Graph.Core.1.*\net45\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Microsoft.Graph.1.*\net45\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
}
elseif ($PSEdition -eq 'Core') {
    foreach ($Path in ($ModuleManifest.FileList -like "*\System.ValueTuple.*\netstandard1.0\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\System.Net.Http.*\netstandard1.3\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Newtonsoft.Json.*\netstandard1.0\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Microsoft.Graph.Core.*\netstandard1.1\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
    foreach ($Path in ($ModuleManifest.FileList -like "*\Microsoft.Graph.*\netstandard1.3\*.dll")) {
        $RequiredAssemblies.Add((Join-Path $PSScriptRoot $Path))
    }
}

## Load correct assemblies for the PowerShell platform
try {
    Add-Type -LiteralPath $RequiredAssemblies | Out-Null
}
catch { throw }
