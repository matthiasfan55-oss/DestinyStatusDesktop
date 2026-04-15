param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir = Join-Path $repoRoot 'src'
$distDir = Join-Path $repoRoot 'dist'
$stageDir = Join-Path $distDir 'package'
$versionPath = Join-Path $repoRoot 'version.json'
$generatedVersionPath = Join-Path $srcDir 'GeneratedVersion.cs'
$wpfDir = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF'

$repoOwner = 'matthiasfan55-oss'
$repoName = 'DestinyStatusDesktop'
$repoBranch = 'main'
$exeName = 'KickStatusSquare.exe'
$packageName = 'KickStatusSquare-package.zip'

if (-not (Test-Path $versionPath)) {
    throw "Missing version file at $versionPath"
}

$versionConfig = Get-Content $versionPath -Raw | ConvertFrom-Json
$version = [string]$versionConfig.version
if ([string]::IsNullOrWhiteSpace($version)) {
    throw 'version.json must contain a non-empty "version" value.'
}

@"
namespace KickStatusApp
{
    internal static class GeneratedVersion
    {
        internal const string Number = "$version";
    }
}
"@ | Set-Content -Path $generatedVersionPath -Encoding UTF8

if (Test-Path $distDir) {
    Get-ChildItem -LiteralPath $distDir -Force | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    }
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) {
    throw "CSC compiler not found at $csc"
}

& $csc `
    /nologo `
    /target:winexe `
    /win32icon:"$srcDir\KickStatusApp.ico" `
    /out:"$stageDir\$exeName" `
    /reference:System.Windows.Forms.dll `
    /reference:System.Drawing.dll `
    /reference:System.Web.Extensions.dll `
    /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Xaml.dll" `
    /reference:"$wpfDir\PresentationCore.dll" `
    /reference:"$wpfDir\WindowsBase.dll" `
    "$srcDir\KickStatusApp.cs" `
    "$generatedVersionPath"

if ($LASTEXITCODE -ne 0) {
    throw "Compilation failed with exit code $LASTEXITCODE"
}

Copy-Item "$srcDir\EditKickStatusConfig.ps1" (Join-Path $stageDir 'EditKickStatusConfig.ps1') -Force
Copy-Item "$srcDir\KickStatusApp.ico" (Join-Path $stageDir 'KickStatusApp.ico') -Force
Copy-Item "$srcDir\config.json" (Join-Path $stageDir 'config.json') -Force
Copy-Item (Join-Path $repoRoot 'Edit Kick Status Channels.cmd') (Join-Path $stageDir 'Edit Kick Status Channels.cmd') -Force
Copy-Item (Join-Path $repoRoot 'Edit Kick Status Channels.vbs') (Join-Path $stageDir 'Edit Kick Status Channels.vbs') -Force

Copy-Item (Join-Path $stageDir $exeName) (Join-Path $distDir $exeName) -Force
Copy-Item (Join-Path $stageDir 'EditKickStatusConfig.ps1') (Join-Path $distDir 'EditKickStatusConfig.ps1') -Force
Copy-Item (Join-Path $stageDir 'Edit Kick Status Channels.cmd') (Join-Path $distDir 'Edit Kick Status Channels.cmd') -Force
Copy-Item (Join-Path $stageDir 'Edit Kick Status Channels.vbs') (Join-Path $distDir 'Edit Kick Status Channels.vbs') -Force
Copy-Item (Join-Path $stageDir 'KickStatusApp.ico') (Join-Path $distDir 'KickStatusApp.ico') -Force
Copy-Item (Join-Path $stageDir 'config.json') (Join-Path $distDir 'config.json') -Force

$packagePath = Join-Path $distDir $packageName
if (Test-Path $packagePath) {
    Remove-Item -LiteralPath $packagePath -Force
}
Compress-Archive -Path (Join-Path $stageDir '*') -DestinationPath $packagePath -Force

$packageUrl = "https://raw.githubusercontent.com/$repoOwner/$repoName/$repoBranch/dist/$packageName"
$updateManifest = [ordered]@{
    version = $version
    packageUrl = $packageUrl
    notes = "Kick Status $version"
} | ConvertTo-Json

$updateManifest | Set-Content -Path (Join-Path $repoRoot 'update.json') -Encoding UTF8

Write-Host "Built $exeName $version"
Write-Host "Package: $packagePath"
Write-Host "Manifest: $(Join-Path $repoRoot 'update.json')"
