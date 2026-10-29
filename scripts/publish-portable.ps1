param(
    [ValidateSet('win-x64')][string]$Runtime = 'win-x64',
    [string]$NuGetConfig,
    [string]$PackageCache,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$repoDirectory = Split-Path -Parent $PSScriptRoot
$runnerDirectory = Join-Path $repoDirectory 'artifacts/runner'
$appDirectory = if ($OutputDirectory) { [System.IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path $repoDirectory 'artifacts/portable' }
Push-Location $repoDirectory
try {
    $restoreArguments = @()
    if ($NuGetConfig) { $restoreArguments += @('--configfile', [System.IO.Path]::GetFullPath($NuGetConfig)) }
    if ($PackageCache) { $restoreArguments += @('--packages', [System.IO.Path]::GetFullPath($PackageCache)) }
    dotnet publish SoftwareInstaller.Runner/SoftwareInstaller.Runner.csproj -c Release -r $Runtime --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -p:DebugSymbols=false --disable-build-servers -m:1 -nodeReuse:false -o $runnerDirectory @restoreArguments
    if ($LASTEXITCODE -ne 0) { throw '执行器发布失败。' }
    dotnet publish SoftwareInstaller.UI/SoftwareInstaller.UI.csproj -c Release -r $Runtime --self-contained true -p:DebugType=None -p:DebugSymbols=false --disable-build-servers -m:1 -nodeReuse:false -o $appDirectory @restoreArguments
    if ($LASTEXITCODE -ne 0) { throw '管理程序发布失败。' }
    Write-Output "便携版已生成：$appDirectory"
} finally { Pop-Location }
