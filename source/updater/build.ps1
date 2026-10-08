$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$payload = Join-Path $projectRoot 'release-1.9.4\Flintec_ControlCenter_1.9.4_Portable.exe'
$outputDir = Join-Path $projectRoot 'release-1.9.4-autoupdate'
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$sourceFile = Join-Path $PSScriptRoot 'Launcher.cs'
$arguments = @('/nologo', '/optimize+', '/platform:anycpu', '/reference:System.Web.Extensions.dll', '/reference:System.Windows.Forms.dll', "/resource:$payload,FlintecApp", $sourceFile)
$appFile = Join-Path $outputDir 'Flintec_ControlCenter_1.9.4_AutoUpdate.exe'
& $compiler '/target:winexe' "/out:$appFile" @arguments
if ($LASTEXITCODE -ne 0) { throw 'Launcher build failed' }
$testFile = Join-Path $outputDir 'Updater.Tests.exe'
& $compiler '/target:exe' "/out:$testFile" @arguments
if ($LASTEXITCODE -ne 0) { throw 'Test build failed' }
& $testFile '--self-test'
if ($LASTEXITCODE -ne 0) { throw 'Updater tests failed' }
Get-FileHash -LiteralPath $appFile -Algorithm SHA256 | ForEach-Object { ($_.Hash.ToLower() + '  ' + [IO.Path]::GetFileName($_.Path)) } | Set-Content -LiteralPath (Join-Path $outputDir 'SHA256SUMS.txt') -Encoding ascii
