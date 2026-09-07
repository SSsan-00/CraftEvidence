[CmdletBinding()]
param(
    [switch]$Publish,
    [ValidateSet('win-x64', 'win-x86')]
    [string]$Runtime = 'win-x64'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$solutionPath = Join-Path $projectRoot 'EvidenceCrafter.sln'
$testProject = Join-Path $projectRoot 'tests\EvidenceCrafter.Tests\EvidenceCrafter.Tests.csproj'
$appProject = Join-Path $projectRoot 'src\EvidenceCrafter.App\EvidenceCrafter.App.csproj'
$publishPath = Join-Path $projectRoot "artifacts\publish\$Runtime"

Push-Location $projectRoot
try {
    $sdkVersion = (& dotnet --version).Trim()
    if (-not $sdkVersion.StartsWith('9.')) {
        throw "EvidenceCrafter requires .NET 9 SDK. Detected: $sdkVersion"
    }

    dotnet restore $solutionPath
    if ($LASTEXITCODE -ne 0) { throw 'dotnet restore failed.' }

    dotnet build $solutionPath --configuration Release --no-restore
    if ($LASTEXITCODE -ne 0) { throw 'dotnet build failed.' }

    dotnet test $testProject --configuration Release --no-build --filter 'TestCategory!=ExcelIntegration'
    if ($LASTEXITCODE -ne 0) { throw 'dotnet test failed.' }

    if ($Publish) {
        $publishRoot = [System.IO.Path]::GetFullPath((Join-Path $projectRoot 'artifacts\publish'))
        $resolvedPublishPath = [System.IO.Path]::GetFullPath($publishPath)
        $requiredPrefix = $publishRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
        if (-not $resolvedPublishPath.StartsWith($requiredPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to clean a publish path outside the repository artifacts directory: $resolvedPublishPath"
        }

        if (Test-Path -LiteralPath $resolvedPublishPath) {
            Remove-Item -LiteralPath $resolvedPublishPath -Recurse -Force
        }

        dotnet publish $appProject `
            --configuration Release `
            --runtime $Runtime `
            --self-contained true `
            --output $resolvedPublishPath `
            -p:PublishSingleFile=true `
            -p:IncludeNativeLibrariesForSelfExtract=true `
            -p:PublishTrimmed=false
        if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed.' }

        $exePath = Join-Path $resolvedPublishPath 'EvidenceCrafter.exe'
        if (-not (Test-Path -LiteralPath $exePath)) {
            throw "Publish completed without the expected executable: $exePath"
        }

        Get-FileHash -Algorithm SHA256 -LiteralPath $exePath |
            ForEach-Object { "{0}  {1}" -f $_.Hash, (Split-Path -Leaf $_.Path) } |
            Set-Content -Encoding ascii -LiteralPath (Join-Path $resolvedPublishPath 'EvidenceCrafter.exe.sha256')

        $allowedPublishFiles = @(
            [System.IO.Path]::GetFullPath($exePath),
            [System.IO.Path]::GetFullPath((Join-Path $resolvedPublishPath 'EvidenceCrafter.exe.sha256'))
        )
        $unexpectedFiles = Get-ChildItem -LiteralPath $resolvedPublishPath -File -Recurse |
            Where-Object { [System.IO.Path]::GetFullPath($_.FullName) -notin $allowedPublishFiles }
        if ($unexpectedFiles) {
            throw "Single-file publish produced unexpected files: $($unexpectedFiles.Name -join ', ')"
        }
    }
}
finally {
    Pop-Location
}
