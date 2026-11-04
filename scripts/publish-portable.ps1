param(
    [ValidateSet('win-x64')][string]$Runtime = 'win-x64',
    [string]$NuGetConfig,
    [string]$PackageCache,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$repoDirectory = Split-Path -Parent $PSScriptRoot
$appDirectory = if ($OutputDirectory) { [System.IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path $repoDirectory 'artifacts/portable' }
Push-Location $repoDirectory
try {
    $restoreArguments = @()
    if ($NuGetConfig) { $restoreArguments += @('--configfile', [System.IO.Path]::GetFullPath($NuGetConfig)) }
    if ($PackageCache) { $restoreArguments += @('--packages', [System.IO.Path]::GetFullPath($PackageCache)) }
    & (Join-Path $PSScriptRoot 'build-runner.ps1') -NuGetConfig $NuGetConfig -PackageCache $PackageCache
    dotnet publish SoftwareInstaller.UI/SoftwareInstaller.UI.csproj -c Release -r $Runtime --self-contained true -p:DebugType=None -p:DebugSymbols=false --disable-build-servers -m:1 -nodeReuse:false -o $appDirectory @restoreArguments
    if ($LASTEXITCODE -ne 0) { throw '管理程序发布失败。' }
    $entryPoint = Join-Path $appDirectory '软件安装助理.exe'
    if (-not (Test-Path -LiteralPath $entryPoint -PathType Leaf)) { throw '未找到软件安装助理.exe，请检查发布结果。' }
    # Remove only the obsolete generated launcher files when reusing an existing output directory.
    foreach ($legacyName in @('SoftwareInstaller.UI.exe', 'SoftwareInstaller.UI.dll', 'SoftwareInstaller.UI.deps.json', 'SoftwareInstaller.UI.runtimeconfig.json', 'SoftwareInstaller.UI.pdb')) {
        $legacyPath = Join-Path $appDirectory $legacyName
        if (Test-Path -LiteralPath $legacyPath -PathType Leaf) { Remove-Item -LiteralPath $legacyPath -Force }
    }
    Write-Output "便携版已生成：$appDirectory"
    Write-Output "启动程序：$entryPoint"
} finally { Pop-Location }
