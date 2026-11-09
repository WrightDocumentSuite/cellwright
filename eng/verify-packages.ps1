[CmdletBinding()]
param(
    [string] $PackageVersion,
    [string] $BaseVersion = '0.1.5',
    [string] $BaselineVersion = '0.1.4',
    [string] $InkwrightVersion = '0.1.5',
    [string] $RuntimeIdentifier,
    [switch] $AllowStablePackageVersion,
    [switch] $SkipSourceLinkContentCheck,
    [switch] $SkipAot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-DotNet {
    param([Parameter(Mandatory)][string[]] $Arguments)

    Write-Host "dotnet $($Arguments -join ' ')"
    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet exited with code $LASTEXITCODE."
    }
}

function Read-ZipText {
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive] $Archive,
        [Parameter(Mandatory)][string] $EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "Archive entry '$EntryName' was not found."
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Get-RequiredDependencyIds {
    param([Parameter(Mandatory)][string] $PackageId)

    switch ($PackageId) {
        'Cellwright' { return @() }
        'Cellwright.Formulas' { return @('Cellwright') }
        'Cellwright.Mapping' { return @('Cellwright') }
        'Cellwright.Pdf' { return @('Cellwright', 'Inkwright') }
        'Cellwright.Xls' { return @('Cellwright', 'Cellwright.Formulas') }
        default { throw "Unexpected package id '$PackageId'." }
    }
}

function Test-Package {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][string] $ExpectedId,
        [Parameter(Mandatory)][string] $ExpectedVersion,
        [Parameter(Mandatory)][string] $ExpectedInkwrightVersion,
        [Parameter(Mandatory)][string] $ExpectedCommit
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $nuspecEntries = @($archive.Entries | Where-Object FullName -Like '*.nuspec')
        if ($nuspecEntries.Count -ne 1) {
            throw "$Path must contain exactly one nuspec; found $($nuspecEntries.Count)."
        }

        [xml] $nuspec = Read-ZipText -Archive $archive -EntryName $nuspecEntries[0].FullName
        $namespace = [System.Xml.XmlNamespaceManager]::new($nuspec.NameTable)
        $namespace.AddNamespace('n', $nuspec.DocumentElement.NamespaceURI)
        $metadata = $nuspec.SelectSingleNode('/n:package/n:metadata', $namespace)
        if ($null -eq $metadata) {
            throw "$Path has no NuGet metadata element."
        }

        $actualId = $metadata.SelectSingleNode('n:id', $namespace).InnerText
        $actualVersion = $metadata.SelectSingleNode('n:version', $namespace).InnerText
        if ($actualId -cne $ExpectedId -or $actualVersion -cne $ExpectedVersion) {
            throw "$Path identifies '$actualId $actualVersion', expected '$ExpectedId $ExpectedVersion'."
        }

        $repository = $metadata.SelectSingleNode('n:repository', $namespace)
        if ($null -eq $repository -or
            $repository.GetAttribute('type') -cne 'git' -or
            $repository.GetAttribute('url') -cne 'https://github.com/WrightDocumentSuite/cellwright.git' -or
            $repository.GetAttribute('commit') -cne $ExpectedCommit) {
            throw "$Path is missing reproducible GitHub SourceLink repository metadata."
        }

        $dependencies = @($metadata.SelectNodes('.//n:dependency', $namespace))
        $requiredDependencyIds = @(Get-RequiredDependencyIds -PackageId $ExpectedId)
        foreach ($requiredId in $requiredDependencyIds) {
            if (-not ($dependencies | Where-Object { $_.GetAttribute('id') -ceq $requiredId })) {
                throw "$Path is missing required dependency '$requiredId'."
            }
        }

        foreach ($dependency in $dependencies) {
            $dependencyId = $dependency.GetAttribute('id')
            if ($dependencyId -match '^Cellwright(?:\.|$)') {
                $expectedRange = "[$ExpectedVersion]"
            }
            elseif ($dependencyId -match '^Inkwright(?:\.|$)') {
                $expectedRange = "[$ExpectedInkwrightVersion]"
            }
            else {
                continue
            }

            $actualRange = $dependency.GetAttribute('version')
            if ($actualRange -cne $expectedRange) {
                throw "$Path dependency '$dependencyId' has '$actualRange'; expected exact range '$expectedRange'."
            }
        }

        if ($ExpectedId -ceq 'Cellwright.Mapping') {
            $analyzers = @($archive.Entries | Where-Object FullName -Like 'analyzers/dotnet/cs/*.dll')
            if ($analyzers.Count -ne 1 -or
                $analyzers[0].FullName -cne 'analyzers/dotnet/cs/Cellwright.Mapping.Generator.dll') {
                throw "$Path must contain exactly the Cellwright.Mapping.Generator analyzer asset."
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    $symbolPath = [System.IO.Path]::ChangeExtension($Path, '.snupkg')
    if (-not [System.IO.File]::Exists($symbolPath)) {
        throw "Symbol package '$symbolPath' was not produced."
    }

    $symbols = [System.IO.Compression.ZipFile]::OpenRead($symbolPath)
    try {
        $expectedPdb = "lib/net10.0/$ExpectedId.pdb"
        if ($null -eq $symbols.GetEntry($expectedPdb)) {
            throw "$symbolPath does not contain '$expectedPdb'."
        }
    }
    finally {
        $symbols.Dispose()
    }
}

function Get-DefaultRuntimeIdentifier {
    $architecture = switch ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture) {
        ([System.Runtime.InteropServices.Architecture]::X64) { 'x64' }
        ([System.Runtime.InteropServices.Architecture]::Arm64) { 'arm64' }
        default { throw "Unsupported AOT architecture '$([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture)'." }
    }

    if ($IsWindows) { return "win-$architecture" }
    if ($IsLinux) { return "linux-$architecture" }
    if ($IsMacOS) { return "osx-$architecture" }
    throw 'Unsupported operating system for the Native AOT package gate.'
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$fullSha = (& git -C $repoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $fullSha -notmatch '^[0-9a-f]{40}$') {
    throw 'Could not determine the Git commit for the package version.'
}
$shortSha = $fullSha.Substring(0, 12)

if (-not $SkipSourceLinkContentCheck) {
    $dirty = @(& git -C $repoRoot status --porcelain=v1 --untracked-files=all)
    if ($LASTEXITCODE -ne 0) {
        throw 'Could not inspect the Git worktree before SourceLink validation.'
    }
    if ($dirty.Count -ne 0) {
        throw 'Strict SourceLink validation requires a clean checkout. Use -SkipSourceLinkContentCheck only for a local diagnostic run.'
    }

    if ($IsWindows) {
        $autoCrlf = @(& git -C $repoRoot config --bool --get core.autocrlf)
        if ($LASTEXITCODE -notin @(0, 1)) {
            throw 'Could not inspect Git core.autocrlf.'
        }
        if ($autoCrlf.Count -ne 0 -and $autoCrlf[0] -cne 'false') {
            throw "Strict SourceLink validation requires core.autocrlf=false; found '$($autoCrlf[0])'."
        }

        $badEndings = @(& git -C $repoRoot ls-files --eol | Where-Object { $_ -match 'w/(?:crlf|mixed)' })
        if ($LASTEXITCODE -ne 0) {
            throw 'Could not inspect Git working-tree line endings.'
        }
        if ($badEndings.Count -ne 0) {
            throw "Strict SourceLink validation rejects CRLF or mixed tracked files:`n$($badEndings -join "`n")"
        }
    }
}

$packageVersionWasSpecified = -not [string]::IsNullOrWhiteSpace($PackageVersion)
$invocationStamp = [DateTime]::UtcNow.ToString("yyyyMMdd't'HHmmssfff", [Globalization.CultureInfo]::InvariantCulture)
$invocationRandom = [Guid]::NewGuid().ToString('N').Substring(0, 8)
if (-not $packageVersionWasSpecified) {
    if ($AllowStablePackageVersion) {
        throw '-AllowStablePackageVersion requires an explicit stable -PackageVersion X.Y.Z.'
    }

    $PackageVersion = "$BaseVersion-dev.$invocationStamp.g$shortSha.$invocationRandom"
}

$isPrerelease = $PackageVersion -match '^\d+\.\d+\.\d+-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*$'
$isStable = $PackageVersion -match '^\d+\.\d+\.\d+$'
if (-not $isPrerelease -and -not $isStable) {
    throw "PackageVersion '$PackageVersion' must be a three-part SemVer version."
}
if ($isStable -and -not $AllowStablePackageVersion) {
    throw "Stable PackageVersion '$PackageVersion' requires -AllowStablePackageVersion."
}
if ($isPrerelease -and $AllowStablePackageVersion) {
    throw '-AllowStablePackageVersion may only be used with a stable -PackageVersion X.Y.Z.'
}
if ($isStable -and $SkipSourceLinkContentCheck) {
    throw 'A stable package cannot skip strict SourceLink content validation.'
}

if ([string]::IsNullOrWhiteSpace($RuntimeIdentifier)) {
    $RuntimeIdentifier = Get-DefaultRuntimeIdentifier
}

$verificationRoot = Join-Path $repoRoot 'artifacts/package-verification'
$runDirectoryName = if ($packageVersionWasSpecified) {
    "$PackageVersion--run.$invocationStamp.$invocationRandom"
}
else {
    $PackageVersion
}
$runRoot = Join-Path $verificationRoot "runs/$runDirectoryName"
$packageDirectory = Join-Path $runRoot 'packages'
$nugetPackages = Join-Path $runRoot 'nuget-packages'
$buildArtifacts = Join-Path $runRoot 'build-artifacts'
$consumerOutput = Join-Path $runRoot "consumer/$RuntimeIdentifier"
$consumerProject = Join-Path $repoRoot 'eng/package-consumer/Cellwright.PackageConsumer.csproj'
$solution = Join-Path $repoRoot 'Cellwright.slnx'
$nugetConfig = Join-Path $repoRoot 'nuget.config'
$verificationLockRoot = Join-Path $runRoot 'locks'

[System.IO.Directory]::CreateDirectory($packageDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory($nugetPackages) | Out-Null
[System.IO.Directory]::CreateDirectory($verificationLockRoot) | Out-Null

$previousNuGetPackages = $env:NUGET_PACKAGES
$env:NUGET_PACKAGES = $nugetPackages
Push-Location $repoRoot
try {
    $allowStableProperty = if ($AllowStablePackageVersion) { 'true' } else { 'false' }
    $commonProperties = @(
        "-p:PackageVersion=$PackageVersion",
        '-p:CellwrightPackageBuild=true',
        "-p:CellwrightAllowStablePackageVersion=$allowStableProperty",
        '-p:CellwrightEnablePackageValidation=true',
        "-p:CellwrightPackageBaselineVersion=$BaselineVersion",
        '-p:RestorePackagesWithLockFile=true',
        '-p:RestoreLockedMode=false',
        "-p:CellwrightVerificationLockRoot=$verificationLockRoot",
        "-p:RestorePackagesPath=$nugetPackages",
        "-p:ArtifactsPath=$buildArtifacts"
    )

    Invoke-DotNet -Arguments (@(
        'restore', $solution,
        '--configfile', $nugetConfig,
        '--packages', $nugetPackages,
        '--verbosity', 'minimal'
    ) + $commonProperties)

    Invoke-DotNet -Arguments (@(
        'pack', $solution,
        '--configuration', 'Release',
        '--no-restore',
        '--output', $packageDirectory,
        '--verbosity', 'minimal'
    ) + $commonProperties)

    $packageIds = @(
        'Cellwright',
        'Cellwright.Formulas',
        'Cellwright.Mapping',
        'Cellwright.Pdf',
        'Cellwright.Xls'
    )

    foreach ($packageId in $packageIds) {
        $packagePath = Join-Path $packageDirectory "$packageId.$PackageVersion.nupkg"
        if (-not [System.IO.File]::Exists($packagePath)) {
            throw "Expected package '$packagePath' was not produced."
        }

        Test-Package `
            -Path $packagePath `
            -ExpectedId $packageId `
            -ExpectedVersion $PackageVersion `
            -ExpectedInkwrightVersion $InkwrightVersion `
            -ExpectedCommit $fullSha
    }

    if (-not $SkipSourceLinkContentCheck) {
        $sourceLinkToolRoot = Join-Path $runRoot 'tools/sourcelink'
        $sourceLinkTool = Join-Path $sourceLinkToolRoot $(if ($IsWindows) { 'sourcelink.exe' } else { 'sourcelink' })
        if (-not [System.IO.File]::Exists($sourceLinkTool)) {
            Invoke-DotNet -Arguments @(
                'tool', 'install', 'sourcelink',
                '--version', '3.1.1',
                '--tool-path', $sourceLinkToolRoot,
                '--configfile', $nugetConfig
            )
        }

        $pdbRoot = Join-Path $runRoot 'sourcelink-pdbs'
        [System.IO.Directory]::CreateDirectory($pdbRoot) | Out-Null

        foreach ($packageId in $packageIds) {
            $symbolPath = Join-Path $packageDirectory "$packageId.$PackageVersion.snupkg"
            $symbols = [System.IO.Compression.ZipFile]::OpenRead($symbolPath)
            try {
                $pdbEntry = $symbols.GetEntry("lib/net10.0/$packageId.pdb")
                if ($null -eq $pdbEntry) {
                    throw "$symbolPath has no portable PDB to validate."
                }

                $pdbPath = Join-Path $pdbRoot "$packageId.pdb"
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pdbEntry, $pdbPath, $true)
            }
            finally {
                $symbols.Dispose()
            }

            $sourceLinkJson = (& $sourceLinkTool print-json $pdbPath) -join "`n"
            if ($LASTEXITCODE -ne 0) {
                throw "sourcelink print-json failed for '$pdbPath'."
            }

            $sourceLink = $sourceLinkJson | ConvertFrom-Json
            $mappings = @($sourceLink.documents.PSObject.Properties)
            if ($mappings.Count -eq 0) {
                throw "$pdbPath has no embedded SourceLink mappings."
            }
            foreach ($mapping in $mappings) {
                $url = [string] $mapping.Value
                if (-not $url.StartsWith('https://raw.githubusercontent.com/WrightDocumentSuite/cellwright/', [StringComparison]::Ordinal) -or
                    -not $url.Contains($fullSha, [StringComparison]::OrdinalIgnoreCase)) {
                    throw "$pdbPath contains unexpected SourceLink URL '$url'."
                }
            }

            & $sourceLinkTool test $pdbPath
            if ($LASTEXITCODE -ne 0) {
                throw "sourcelink test failed for '$pdbPath'."
            }
        }
    }

    if (-not $SkipAot) {
        $consumerNuGetConfig = Join-Path $runRoot 'consumer.NuGet.config'
        [xml] $consumerNuGet = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="package-under-test" value="placeholder" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
</configuration>
'@
        $consumerNuGet.configuration.packageSources.add[0].SetAttribute('value', $packageDirectory)
        $consumerNuGet.Save($consumerNuGetConfig)

        $consumerProperties = @(
            "-p:CellwrightPackageVersion=$PackageVersion",
            '-p:RestorePackagesWithLockFile=true',
            '-p:RestoreLockedMode=false',
            "-p:CellwrightVerificationLockRoot=$verificationLockRoot",
            "-p:RestorePackagesPath=$nugetPackages",
            "-p:ArtifactsPath=$buildArtifacts",
            '-p:PublishAot=true',
            '-p:PublishTrimmed=true'
        )

        Invoke-DotNet -Arguments (@(
            'restore', $consumerProject,
            '--runtime', $RuntimeIdentifier,
            '--configfile', $consumerNuGetConfig,
            '--packages', $nugetPackages,
            '--verbosity', 'minimal'
        ) + $consumerProperties)

        Invoke-DotNet -Arguments (@(
            'publish', $consumerProject,
            '--configuration', 'Release',
            '--runtime', $RuntimeIdentifier,
            '--no-restore',
            '--output', $consumerOutput,
            '--verbosity', 'minimal'
        ) + $consumerProperties)

        $executableName = if ($IsWindows) { 'Cellwright.PackageConsumer.exe' } else { 'Cellwright.PackageConsumer' }
        $executable = Join-Path $consumerOutput $executableName
        if (-not [System.IO.File]::Exists($executable)) {
            throw "Native AOT executable '$executable' was not produced."
        }

        & $executable
        if ($LASTEXITCODE -ne 0) {
            throw "Native AOT package consumer exited with code $LASTEXITCODE."
        }
    }

    $manifest = [ordered]@{
        packageVersion = $PackageVersion
        baselineVersion = $BaselineVersion
        inkwrightVersion = $InkwrightVersion
        runtimeIdentifier = $RuntimeIdentifier
        packages = $packageIds
        sourceLinkContent = -not $SkipSourceLinkContentCheck
        nativeAot = -not $SkipAot
        stablePackage = [bool] $AllowStablePackageVersion
    }
    $manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runRoot 'verification.json') -Encoding utf8

    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_OUTPUT)) {
        "package-version=$PackageVersion" | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding utf8
        "package-directory=$packageDirectory" | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding utf8
    }

    Write-Host "Verified Cellwright package set $PackageVersion against $BaselineVersion."
    Write-Host "Packages: $packageDirectory"
}
finally {
    Pop-Location
    $env:NUGET_PACKAGES = $previousNuGetPackages
}
