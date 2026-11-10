param(
    [string]$NuGetConfig,
    [string]$PackageCache,
    [switch]$NoRestore
)
$ErrorActionPreference = 'Stop'
$repoDirectory = Split-Path -Parent $PSScriptRoot
$runnerDirectory = Join-Path $repoDirectory 'artifacts/runner'
Push-Location $repoDirectory
try {
    $restoreArguments = @()
    if (-not $NoRestore) {
        if ($NuGetConfig) { $restoreArguments += @('--configfile', [System.IO.Path]::GetFullPath($NuGetConfig)) }
        if ($PackageCache) { $restoreArguments += @('--packages', [System.IO.Path]::GetFullPath($PackageCache)) }
    }
    if ($NoRestore) { $restoreArguments += '--no-restore' }
    # Bundle the .NET 9 runtime again; keep the current resource, queue and process-control code.
    dotnet publish SoftwareInstaller.Runner/SoftwareInstaller.Runner.csproj -c Release -f net9.0-windows -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=false -p:DebugType=None -p:DebugSymbols=false --disable-build-servers -m:1 -nodeReuse:false -o $runnerDirectory @restoreArguments
    if ($LASTEXITCODE -ne 0) { throw '执行器构建失败。' }
    Write-Output "执行器模板已更新：$runnerDirectory"
} finally { Pop-Location }
