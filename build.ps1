Function Info($msg) {
  Write-Host -ForegroundColor DarkGreen "`nINFO: $msg`n"
}

Function Error($msg) {
  Write-Host `n`n
  Write-Error $msg
  exit 1
}

Function CheckReturnCodeOfPreviousCommand($msg) {
  if(-Not $?) {
    Error "${msg}. Error code: $LastExitCode"
  }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Resolve-Path $PSScriptRoot
$buildDir = "$root/build"

Info "Find Visual Studio installation path"
$vswhereCommand = Get-Command -Name "${Env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$installationPath = & $vswhereCommand -prerelease -latest -property installationPath

Info "Open Visual Studio 2022 Developer PowerShell"
& $installationPath\Common7\Tools\Launch-VsDevShell.ps1 -Arch amd64

Info "Remove '$buildDir' folder if it exists"
Remove-Item $buildDir -Force -Recurse -ErrorAction SilentlyContinue
New-Item $buildDir -Force -ItemType "directory" > $null

Info "Download nuget.exe"
Invoke-WebRequest -Uri https://dist.nuget.org/win-x86-commandline/latest/nuget.exe -OutFile $buildDir/nuget.exe

Info "Download comexplore source code"
Invoke-WebRequest -Uri https://github.com/trieck/comexplore/archive/8efcc77909cff3433dbe306197fab6a4c135a9fa.zip -OutFile $buildDir/comexplore.zip

Info "Extract the source code"
[System.IO.Compression.ZipFile]::ExtractToDirectory("$buildDir/comexplore.zip", $buildDir)
Rename-Item -Path $buildDir/comexplore-8efcc77909cff3433dbe306197fab6a4c135a9fa -NewName comexplore

Info "Patch comexplore source code: set EmbedManifest to true"
(Get-Content $buildDir/comexplore/comexplore.vcxproj) | ForEach-Object { $_ -replace '<EmbedManifest>false</EmbedManifest>', '<EmbedManifest>true</EmbedManifest>' } | Set-Content $buildDir/comexplore/comexplore.vcxproj

Info "Restore NuGet packages"
& $buildDir/nuget.exe restore $buildDir/comexplore/comexplore.sln
CheckReturnCodeOfPreviousCommand "nuget.exe restore failed"

Info "Build project"
msbuild `
  /nologo `
  /verbosity:minimal `
  /property:Configuration=Release `
  /property:Platform=x64 `
  $buildDir/comexplore/comexplore.sln
CheckReturnCodeOfPreviousCommand "build failed"

Info "Create zip archive"
New-Item $buildDir/Publish -Force -ItemType "directory" > $null
Compress-Archive -Path $buildDir/comexplore/x64/Release/comexplore.exe -DestinationPath $buildDir/Publish/comexplore.zip
