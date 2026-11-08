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
    if (-not $NoRestore) {
        dotnet restore SoftwareInstaller.Runner/SoftwareInstaller.Runner.csproj @restoreArguments
        if ($LASTEXITCODE -ne 0) { throw '执行器依赖还原失败。' }
    }
    # Framework runner is a managed EXE with its JSON dependency embedded in memory.
    # It uses the installed Windows 7-compatible CLR; no runtime extraction or compression.
    dotnet build SoftwareInstaller.Runner/SoftwareInstaller.Runner.csproj -c Release -f net48 -p:DebugType=None -p:DebugSymbols=false --no-restore --disable-build-servers -m:1 -nodeReuse:false
    if ($LASTEXITCODE -ne 0) { throw '执行器构建失败。' }
    New-Item -ItemType Directory -Force -Path $runnerDirectory | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoDirectory 'SoftwareInstaller.Runner/bin/Release/net48/SoftwareInstaller.Runner.exe') -Destination $runnerDirectory -Force
    Write-Output "执行器模板已更新：$runnerDirectory"
} finally { Pop-Location }
