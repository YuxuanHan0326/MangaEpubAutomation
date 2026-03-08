<#
.SYNOPSIS
Upscale images inside EPUB files and write new EPUB outputs.

.DESCRIPTION
Supports two mutually exclusive input modes:
1) `-InputEpubPath` for one or many EPUB files.
2) `-InputEpubDirectory` for recursive folder scan.

The script emits machine-readable `PIPELINE_EVENT:` lines when `-GuiMode` is used.

.PARAMETER InputEpubPath
One or many EPUB files to process.

.PARAMETER InputEpubDirectory
Directory to recursively scan for `.epub`.

.PARAMETER OutputSuffix
Suffix appended to output EPUB file name. Default: `-upscaled`.

.PARAMETER NoOutputSuffix
Do not append any suffix to output filename.
Useful for CLI because `powershell -File ... -OutputSuffix ""` may drop empty arguments.

.PARAMETER OutputDirectory
Optional output directory for generated EPUB files.
If not provided, current working directory is used.

.PARAMETER OverwriteOutputEpub
Overwrite existing output EPUB files. Default is skip existing files.

.PARAMETER PlanOnly
Build plan + preflight only, then exit.

.PARAMETER DryRun
Simulate execution without writing output EPUB files.

.PARAMETER AutoConfirm
Skip interactive confirmation.

.PARAMETER GuiMode
Enable `PIPELINE_EVENT` output for GUI subscribers.

.PARAMETER DepsConfigPath
Path to dependency JSON. Default: `<repo>\manga_epub_automation.deps.json`.

.PARAMETER ConfigPath
Path to pipeline config JSON. Default: `<repo>\manga_epub_automation.config.json`.

.PARAMETER UpscaleFactor
Override upscale factor (1..4).

.PARAMETER LossyQuality
Override lossy compression quality (1..100).

.PARAMETER GrayscaleDetectionThreshold
Override grayscale detection threshold (0..24).

.PARAMETER LogLevel
`info` or `debug`.

.PARAMETER FailOnPreflightWarnings
Treat warnings as blocking errors.
#>

[CmdletBinding(DefaultParameterSetName = 'ByFiles')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ByFiles')]
    [string[]]$InputEpubPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByDirectory')]
    [string]$InputEpubDirectory,

    [string]$OutputSuffix = '-upscaled',
    [switch]$NoOutputSuffix,
    [string]$OutputDirectory,
    [switch]$OverwriteOutputEpub,
    [switch]$PlanOnly,
    [switch]$DryRun,
    [switch]$AutoConfirm,
    [switch]$GuiMode,
    [string]$DepsConfigPath,
    [string]$ConfigPath,

    [ValidateSet(1, 2, 3, 4)]
    [int]$UpscaleFactor,

    [ValidateRange(1, 100)]
    [int]$LossyQuality,

    [ValidateRange(0, 24)]
    [int]$GrayscaleDetectionThreshold,

    [ValidateSet('info', 'debug')]
    [string]$LogLevel = 'info',

    [switch]$FailOnPreflightWarnings
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:PipelineGuiMode = [bool]$GuiMode

function Write-Info([string]$Message) { Write-Host "[INFO] $Message" }
function Write-Warn([string]$Message) { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Test-DebugLog { return $LogLevel -eq 'debug' }
function Write-DebugLine([string]$Message) { if (Test-DebugLog) { Write-Host "[DEBUG] $Message" -ForegroundColor DarkGray } }

function Write-GuiEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [Parameter(Mandatory = $true)][object]$Data
    )
    if (-not $script:PipelineGuiMode) { return }
    $payload = [pscustomobject]@{
        type = $Type
        ts_utc = [DateTime]::UtcNow.ToString('o')
        data = $Data
    }
    $json = $payload | ConvertTo-Json -Depth 100 -Compress
    Write-Host ("PIPELINE_EVENT: {0}" -f $json)
}

function Write-JsonUtf8NoBom {
    param(
        [Parameter(Mandatory = $true)][object]$Object,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $json = $Object | ConvertTo-Json -Depth 100
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    }
}

function Convert-ToIntBounded {
    param(
        [object]$Value,
        [int]$DefaultValue,
        [int]$MinValue,
        [int]$MaxValue
    )
    try { $parsed = [int]$Value } catch { return $DefaultValue }
    if ($parsed -lt $MinValue) { return $MinValue }
    if ($parsed -gt $MaxValue) { return $MaxValue }
    return $parsed
}

function Get-DefaultDepsConfigObject {
    param([Parameter(Mandatory = $true)][string]$ScriptRoot)
    $roaming = [Environment]::GetFolderPath('ApplicationData')
    $local = [Environment]::GetFolderPath('LocalApplicationData')
    return [pscustomobject]@{
        version = 1
        paths = [pscustomobject]@{
            python_exe = Join-Path $roaming 'MangaJaNaiConverterGui\python\python\python.exe'
            backend_script = Join-Path $local 'MangaJaNaiConverterGui\current\backend\src\run_upscale.py'
            models_dir = Join-Path $roaming 'MangaJaNaiConverterGui\models'
            kcc_exe = Join-Path $ScriptRoot 'kcc_c2e_9.4.3.exe'
        }
        progress = [pscustomobject]@{
            enabled = $true
            refresh_seconds = 1
            eta_min_samples = 8
            noninteractive_log_interval_seconds = 10
        }
    }
}

function Resolve-DependenciesConfig {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [string]$DepsConfigPath
    )

    $explicitDepsPath = -not [string]::IsNullOrWhiteSpace($DepsConfigPath)
    $resolvedDepsPath = if ($explicitDepsPath) { $DepsConfigPath } else { Join-Path $ScriptRoot 'manga_epub_automation.deps.json' }
    $configExists = Test-Path -LiteralPath $resolvedDepsPath -PathType Leaf

    $depsConfig = $null
    $depsConfigValid = $false
    $depsConfigError = $null
    if ($configExists) {
        try {
            $depsConfig = Read-JsonFile -Path $resolvedDepsPath
            $depsConfigValid = $true
        }
        catch {
            $depsConfigError = $_.Exception.Message
        }
    }

    $defaults = Get-DefaultDepsConfigObject -ScriptRoot $ScriptRoot

    $resolvedPaths = [ordered]@{}
    foreach ($key in @('python_exe', 'backend_script', 'models_dir', 'kcc_exe')) {
        $value = ''
        if ($depsConfigValid -and $depsConfig -and
            ($depsConfig.PSObject.Properties.Name -contains 'paths') -and $depsConfig.paths -and
            ($depsConfig.paths.PSObject.Properties.Name -contains $key)) {
            $value = [string]$depsConfig.paths.$key
        }
        if ([string]::IsNullOrWhiteSpace($value) -and ($defaults.paths.PSObject.Properties.Name -contains $key)) {
            $value = [string]$defaults.paths.$key
        }
        $resolvedPaths[$key] = $value
    }

    $progressEnabled = [bool]$defaults.progress.enabled
    $refreshSeconds = [int]$defaults.progress.refresh_seconds
    $etaMinSamples = [int]$defaults.progress.eta_min_samples
    $nonInteractiveInterval = [int]$defaults.progress.noninteractive_log_interval_seconds

    if ($depsConfigValid -and $depsConfig -and ($depsConfig.PSObject.Properties.Name -contains 'progress') -and $depsConfig.progress) {
        if ($depsConfig.progress.PSObject.Properties.Name -contains 'enabled') {
            $progressEnabled = [bool]$depsConfig.progress.enabled
        }
        $refreshSeconds = Convert-ToIntBounded -Value $depsConfig.progress.refresh_seconds -DefaultValue $refreshSeconds -MinValue 1 -MaxValue 60
        $etaMinSamples = Convert-ToIntBounded -Value $depsConfig.progress.eta_min_samples -DefaultValue $etaMinSamples -MinValue 1 -MaxValue 1000
        $nonInteractiveInterval = Convert-ToIntBounded -Value $depsConfig.progress.noninteractive_log_interval_seconds -DefaultValue $nonInteractiveInterval -MinValue 1 -MaxValue 3600
    }

    return [pscustomobject]@{
        ConfigPath = $resolvedDepsPath
        ConfigPathExplicit = [bool]$explicitDepsPath
        ConfigExists = [bool]$configExists
        ConfigValid = [bool]$depsConfigValid
        ConfigError = $depsConfigError
        ResolvedPaths = [pscustomobject]$resolvedPaths
        ProgressConfig = [pscustomobject]@{
            Enabled = [bool]$progressEnabled
            RefreshSeconds = [int]$refreshSeconds
            EtaMinSamples = [int]$etaMinSamples
            NonInteractiveLogIntervalSeconds = [int]$nonInteractiveInterval
        }
    }
}

function Resolve-PipelineConfig {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [string]$ConfigPath
    )

    $resolvedPath = if ([string]::IsNullOrWhiteSpace($ConfigPath)) { Join-Path $ScriptRoot 'manga_epub_automation.config.json' } else { $ConfigPath }
    $exists = Test-Path -LiteralPath $resolvedPath -PathType Leaf
    if (-not $exists) {
        throw "Config file not found: $resolvedPath"
    }

    $cfg = Read-JsonFile -Path $resolvedPath
    if (-not $cfg) {
        throw "Config parse failed: $resolvedPath"
    }

    $manga = if (($cfg.PSObject.Properties.Name -contains 'manga') -and $cfg.manga) { $cfg.manga } else { [pscustomobject]@{} }
    $workflowOverrides = if (($manga.PSObject.Properties.Name -contains 'WorkflowOverrides') -and $manga.WorkflowOverrides) { $manga.WorkflowOverrides } else { [pscustomobject]@{} }

    $resolved = [pscustomobject]@{
        Path = $resolvedPath
        RawConfig = $cfg
        Manga = [pscustomobject]@{
            OutputFormat = if (($manga.PSObject.Properties.Name -contains 'OutputFormat') -and $manga.OutputFormat) { [string]$manga.OutputFormat } else { 'webp' }
            UpscaleScaleFactor = if ($manga.PSObject.Properties.Name -contains 'UpscaleScaleFactor') { [int]$manga.UpscaleScaleFactor } else { 2 }
            LossyCompressionQuality = if ($manga.PSObject.Properties.Name -contains 'LossyCompressionQuality') { [int]$manga.LossyCompressionQuality } else { 80 }
            GrayscaleDetectionThreshold = if ($workflowOverrides.PSObject.Properties.Name -contains 'GrayscaleDetectionThreshold') { [int]$workflowOverrides.GrayscaleDetectionThreshold } else { 12 }
        }
    }
    return $resolved
}

function Get-WorkflowRef {
    param([Parameter(Mandatory = $true)][object]$Config)
    if (-not ($Config.PSObject.Properties.Name -contains 'Workflows')) { throw 'Invalid settings: missing Workflows' }
    if (-not ($Config.Workflows.PSObject.Properties.Name -contains '$values')) { throw "Invalid settings: Workflows missing `'$values`'" }
    $workflowValues = @($Config.Workflows.'$values')
    if ($workflowValues.Count -lt 1) { throw 'Invalid settings: Workflows.$values is empty' }
    if (-not ($Config.PSObject.Properties.Name -contains 'SelectedWorkflowIndex')) {
        $Config | Add-Member -NotePropertyName SelectedWorkflowIndex -NotePropertyValue 0
    }
    $wfIndex = [int]$Config.SelectedWorkflowIndex
    if ($wfIndex -lt 0 -or $wfIndex -ge $workflowValues.Count) {
        $wfIndex = 0
        $Config.SelectedWorkflowIndex = 0
    }
    return $workflowValues[$wfIndex]
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )
    $baseFull = [System.IO.Path]::GetFullPath($BasePath)
    if (-not $baseFull.EndsWith('\')) { $baseFull += '\' }
    $targetFull = [System.IO.Path]::GetFullPath($TargetPath)
    $baseUri = [System.Uri]::new($baseFull)
    $targetUri = [System.Uri]::new($targetFull)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('/', '\')
}

function Convert-ToPosixPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    return $Path.Replace('\', '/')
}

function Convert-ToUriPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    $segments = Convert-ToPosixPath -Path $Path -split '/'
    return (($segments | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/')
}

function Get-MimeTypeByExtension {
    param([Parameter(Mandatory = $true)][string]$Extension)
    switch ($Extension.ToLowerInvariant()) {
        '.jpg' { return 'image/jpeg' }
        '.jpeg' { return 'image/jpeg' }
        '.png' { return 'image/png' }
        '.webp' { return 'image/webp' }
        '.avif' { return 'image/avif' }
        '.bmp' { return 'image/bmp' }
        default { return 'application/octet-stream' }
    }
}

function New-PreflightIssue {
    param(
        [Parameter(Mandatory = $true)][string]$Severity,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message
    )
    return [pscustomobject]@{
        Severity = $Severity
        Code = $Code
        Message = $Message
    }
}

function Get-InputEpubFiles {
    param(
        [Parameter(Mandatory = $true)][string]$ParameterSetName,
        [string[]]$InputEpubPath,
        [string]$InputEpubDirectory
    )
    $all = New-Object 'System.Collections.Generic.List[string]'
    if ($ParameterSetName -eq 'ByFiles') {
        $rawItems = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in @($InputEpubPath)) {
            if ([string]::IsNullOrWhiteSpace($item)) { continue }
            $text = [string]$item
            $parts = @($text -split "(?:`r`n|`n|`r|\|)")
            if ($parts.Count -lt 2) {
                $rawItems.Add($text) | Out-Null
                continue
            }
            foreach ($part in $parts) {
                if ([string]::IsNullOrWhiteSpace($part)) { continue }
                $rawItems.Add($part.Trim()) | Out-Null
            }
        }

        foreach ($item in $rawItems) {
            try {
                $full = [System.IO.Path]::GetFullPath($item)
                $all.Add($full) | Out-Null
            }
            catch {
                # keep invalid path for preflight reporting
                $all.Add([string]$item) | Out-Null
            }
        }
    }
    else {
        $dirFull = [System.IO.Path]::GetFullPath($InputEpubDirectory)
        foreach ($file in @(Get-ChildItem -LiteralPath $dirFull -Recurse -File -Filter '*.epub' -ErrorAction SilentlyContinue)) {
            $all.Add($file.FullName) | Out-Null
        }
    }

    $unique = @{}
    $deduped = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $all) {
        if (-not $unique.ContainsKey($p)) {
            $unique[$p] = $true
            $deduped.Add($p) | Out-Null
        }
    }
    return @($deduped | Sort-Object)
}

function Convert-ToObjectArray {
    param([object]$Value)

    if ($null -eq $Value) {
        return [object[]]@()
    }

    if ($Value -is [System.Array]) {
        return [object[]]$Value
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        $list = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Value) {
            $list.Add($item) | Out-Null
        }
        return [object[]]$list.ToArray()
    }

    return [object[]]@($Value)
}

function Test-ObjectHasProperty {
    param(
        [object]$Object,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    if ($null -eq $Object) {
        return $false
    }

    try {
        return $null -ne $Object.PSObject.Properties[$PropertyName]
    }
    catch {
        return $false
    }
}

function Get-EpubImageEntries {
    param([Parameter(Mandatory = $true)][string]$ExtractRoot)

    $containerPath = Join-Path $ExtractRoot 'META-INF\container.xml'
    $opfPath = $null
    if (Test-Path -LiteralPath $containerPath -PathType Leaf) {
        try {
            [xml]$containerXml = Get-Content -LiteralPath $containerPath -Raw -Encoding UTF8
            $rootFileNode = $containerXml.SelectSingleNode("/*[local-name()='container']/*[local-name()='rootfiles']/*[local-name()='rootfile']")
            if ($null -ne $rootFileNode) {
                $opfRel = [string]$rootFileNode.GetAttribute('full-path')
                if (-not [string]::IsNullOrWhiteSpace($opfRel)) {
                    $candidate = Join-Path $ExtractRoot ($opfRel -replace '/', '\')
                    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                        $opfPath = $candidate
                    }
                }
            }
        }
        catch {
            Write-DebugLine ("Failed to parse container.xml: {0}" -f $_.Exception.Message)
        }
    }

    if (-not $opfPath) {
        $fallbackOpf = @(Get-ChildItem -LiteralPath $ExtractRoot -Recurse -File -Filter '*.opf' -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($fallbackOpf.Count -gt 0) { $opfPath = $fallbackOpf[0].FullName }
    }

    $entries = New-Object 'System.Collections.Generic.List[object]'
    if ($opfPath) {
        try {
            $opfDoc = New-Object System.Xml.XmlDocument
            $opfDoc.PreserveWhitespace = $true
            $opfDoc.Load($opfPath)
            $opfDir = Split-Path -Parent $opfPath
            $ns = New-Object System.Xml.XmlNamespaceManager($opfDoc.NameTable)
            if ($opfDoc.DocumentElement -and -not [string]::IsNullOrWhiteSpace($opfDoc.DocumentElement.NamespaceURI)) {
                $ns.AddNamespace('opf', $opfDoc.DocumentElement.NamespaceURI)
            }

            $manifestNodes = @()
            if ($ns.HasNamespace('opf')) {
                $manifestNodes = @($opfDoc.SelectNodes('//opf:manifest/opf:item', $ns))
            }
            else {
                $manifestNodes = @($opfDoc.SelectNodes('//manifest/item'))
            }

            foreach ($node in $manifestNodes) {
                $mediaType = [string]$node.GetAttribute('media-type')
                if ([string]::IsNullOrWhiteSpace($mediaType) -or -not $mediaType.StartsWith('image/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                $href = [string]$node.GetAttribute('href')
                if ([string]::IsNullOrWhiteSpace($href)) { continue }
                $resolved = Join-Path $opfDir ($href -replace '/', '\')
                try { $resolved = [System.IO.Path]::GetFullPath($resolved) } catch { continue }
                if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { continue }
                $rel = Get-RelativePath -BasePath $ExtractRoot -TargetPath $resolved
                $entries.Add([pscustomobject]@{
                        FullPath = $resolved
                        RelPath = $rel
                        SourceExtension = [System.IO.Path]::GetExtension($resolved).ToLowerInvariant()
                    }) | Out-Null
            }
        }
        catch {
            Write-DebugLine ("Failed to parse OPF image manifest: {0}" -f $_.Exception.Message)
        }
    }

    if ($entries.Count -lt 1) {
        $fallbackExt = @('.jpg', '.jpeg', '.png', '.webp', '.avif', '.bmp')
        foreach ($file in @(Get-ChildItem -LiteralPath $ExtractRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $fallbackExt -contains $_.Extension.ToLowerInvariant() })) {
            $rel = Get-RelativePath -BasePath $ExtractRoot -TargetPath $file.FullName
            $entries.Add([pscustomobject]@{
                    FullPath = $file.FullName
                    RelPath = $rel
                    SourceExtension = $file.Extension.ToLowerInvariant()
                }) | Out-Null
        }
    }

    $dedup = @{}
    $uniqueEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in $entries) {
        $relPath = [string]$entry.RelPath
        if (-not $dedup.ContainsKey($relPath)) {
            $dedup[$relPath] = $true
            $uniqueEntries.Add($entry) | Out-Null
        }
    }

    return [pscustomobject]@{
        OpfPath = $opfPath
        Entries = @($uniqueEntries | Sort-Object RelPath)
    }
}

function Get-UpscaleFormatForExtension {
    param([Parameter(Mandatory = $true)][string]$Extension)
    switch ($Extension.ToLowerInvariant()) {
        '.jpg' { return 'jpeg' }
        '.jpeg' { return 'jpeg' }
        '.png' { return 'png' }
        '.webp' { return 'webp' }
        '.avif' { return 'avif' }
        '.bmp' { return 'png' }
        default { return 'png' }
    }
}

function Get-ProducedImagePath {
    param(
        [Parameter(Mandatory = $true)][string]$GroupOutputRoot,
        [Parameter(Mandatory = $true)][string]$RelPath,
        [Parameter(Mandatory = $true)][string]$Format
    )
    $dirRel = Split-Path -Parent $RelPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($RelPath)
    $searchDir = if ([string]::IsNullOrWhiteSpace($dirRel)) { $GroupOutputRoot } else { Join-Path $GroupOutputRoot $dirRel }
    if (-not (Test-Path -LiteralPath $searchDir -PathType Container)) {
        return $null
    }

    $candidates = New-Object 'System.Collections.Generic.List[string]'
    switch ($Format.ToLowerInvariant()) {
        'jpeg' {
            $candidates.Add((Join-Path $searchDir ($baseName + '.jpeg'))) | Out-Null
            $candidates.Add((Join-Path $searchDir ($baseName + '.jpg'))) | Out-Null
        }
        default {
            $candidates.Add((Join-Path $searchDir ($baseName + '.' + $Format.ToLowerInvariant()))) | Out-Null
        }
    }

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return $path
        }
    }

    $fallback = @(Get-ChildItem -LiteralPath $searchDir -File -ErrorAction SilentlyContinue | Where-Object {
            [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $baseName
        } | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1)
    if ($fallback.Count -gt 0) { return $fallback[0].FullName }
    return $null
}

function Update-XmlManifestReferences {
    param(
        [Parameter(Mandatory = $true)][string]$OpfPath,
        [Parameter(Mandatory = $true)][string]$ExtractRoot,
        [Parameter(Mandatory = $true)][hashtable]$RenameMap
    )
    if ($RenameMap.Count -lt 1) { return }
    if (-not (Test-Path -LiteralPath $OpfPath -PathType Leaf)) { return }

    $opfDoc = New-Object System.Xml.XmlDocument
    $opfDoc.PreserveWhitespace = $true
    $opfDoc.Load($OpfPath)
    $opfDir = Split-Path -Parent $OpfPath

    $ns = New-Object System.Xml.XmlNamespaceManager($opfDoc.NameTable)
    if ($opfDoc.DocumentElement -and -not [string]::IsNullOrWhiteSpace($opfDoc.DocumentElement.NamespaceURI)) {
        $ns.AddNamespace('opf', $opfDoc.DocumentElement.NamespaceURI)
    }
    $manifestNodes = @()
    if ($ns.HasNamespace('opf')) {
        $manifestNodes = @($opfDoc.SelectNodes('//opf:manifest/opf:item', $ns))
    }
    else {
        $manifestNodes = @($opfDoc.SelectNodes('//manifest/item'))
    }

    $changed = $false
    foreach ($node in $manifestNodes) {
        $href = [string]$node.GetAttribute('href')
        if ([string]::IsNullOrWhiteSpace($href)) { continue }
        $resolvedOld = [System.IO.Path]::GetFullPath((Join-Path $opfDir ($href -replace '/', '\')))
        $relOld = Get-RelativePath -BasePath $ExtractRoot -TargetPath $resolvedOld
        if ($RenameMap.ContainsKey($relOld)) {
            $relNew = [string]$RenameMap[$relOld]
            $absNew = [System.IO.Path]::GetFullPath((Join-Path $ExtractRoot $relNew))
            $hrefNew = Convert-ToPosixPath -Path (Get-RelativePath -BasePath $opfDir -TargetPath $absNew)
            $node.SetAttribute('href', $hrefNew) | Out-Null
            $node.SetAttribute('media-type', (Get-MimeTypeByExtension -Extension ([System.IO.Path]::GetExtension($relNew)))) | Out-Null
            $changed = $true
        }
    }

    if ($changed) {
        $writerSettings = New-Object System.Xml.XmlWriterSettings
        $writerSettings.Indent = $false
        $writerSettings.Encoding = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.Xml.XmlWriter]::Create($OpfPath, $writerSettings)
        try { $opfDoc.Save($writer) } finally { $writer.Dispose() }
    }
}

function Update-TextReferenceFiles {
    param(
        [Parameter(Mandatory = $true)][string]$ExtractRoot,
        [Parameter(Mandatory = $true)][hashtable]$RenameMap
    )
    if ($RenameMap.Count -lt 1) { return }
    $textExt = @('.xhtml', '.html', '.htm', '.xml', '.css', '.ncx', '.svg', '.opf')
    $files = @(Get-ChildItem -LiteralPath $ExtractRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
            $textExt -contains $_.Extension.ToLowerInvariant()
        })

    foreach ($file in $files) {
        $raw = $null
        try {
            $raw = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
        }
        catch {
            try { $raw = Get-Content -LiteralPath $file.FullName -Raw } catch { $raw = $null }
        }
        if ($null -eq $raw) { continue }
        $updated = $raw

        foreach ($key in @($RenameMap.Keys | Sort-Object { $_.Length } -Descending)) {
            $oldRelPosix = Convert-ToPosixPath -Path ([string]$key)
            $newRelPosix = Convert-ToPosixPath -Path ([string]$RenameMap[$key])
            $oldUri = Convert-ToUriPath -Path $oldRelPosix
            $newUri = Convert-ToUriPath -Path $newRelPosix

            if ($updated.Contains($oldRelPosix)) { $updated = $updated.Replace($oldRelPosix, $newRelPosix) }
            if ($updated.Contains($oldUri)) { $updated = $updated.Replace($oldUri, $newUri) }
        }

        if (-not [string]::Equals($raw, $updated, [System.StringComparison]::Ordinal)) {
            [System.IO.File]::WriteAllText($file.FullName, $updated, [System.Text.UTF8Encoding]::new($false))
        }
    }
}

function New-RuntimeSettingsFile {
    param(
        [Parameter(Mandatory = $true)][object]$BaseSettings,
        [Parameter(Mandatory = $true)][string]$RuntimePath,
        [Parameter(Mandatory = $true)][string]$InputFolderPath,
        [Parameter(Mandatory = $true)][string]$OutputFolderPath,
        [Parameter(Mandatory = $true)][string]$ModelsDirectory,
        [Parameter(Mandatory = $true)][int]$ScaleFactor,
        [Parameter(Mandatory = $true)][int]$LossyQuality,
        [Parameter(Mandatory = $true)][int]$GrayThreshold,
        [Parameter(Mandatory = $true)][string]$OutputFormat
    )
    $cfg = ($BaseSettings | ConvertTo-Json -Depth 100 | ConvertFrom-Json)
    if ($cfg.PSObject.Properties.Name -contains 'ModelsDirectory') { $cfg.ModelsDirectory = $ModelsDirectory }
    else { $cfg | Add-Member -Force -NotePropertyName ModelsDirectory -NotePropertyValue $ModelsDirectory }

    $wf = Get-WorkflowRef -Config $cfg

    if ($wf.PSObject.Properties.Name -contains 'SelectedTabIndex') { $wf.SelectedTabIndex = 1 }
    else { $wf | Add-Member -Force -NotePropertyName SelectedTabIndex -NotePropertyValue 1 }

    if ($wf.PSObject.Properties.Name -contains 'InputFilePath') { $wf.InputFilePath = '' }
    if ($wf.PSObject.Properties.Name -contains 'InputFolderPath') { $wf.InputFolderPath = $InputFolderPath }
    else { $wf | Add-Member -Force -NotePropertyName InputFolderPath -NotePropertyValue $InputFolderPath }
    if ($wf.PSObject.Properties.Name -contains 'OutputFolderPath') { $wf.OutputFolderPath = $OutputFolderPath }
    else { $wf | Add-Member -Force -NotePropertyName OutputFolderPath -NotePropertyValue $OutputFolderPath }
    if ($wf.PSObject.Properties.Name -contains 'OutputFilename') { $wf.OutputFilename = '%filename%' }
    else { $wf | Add-Member -Force -NotePropertyName OutputFilename -NotePropertyValue '%filename%' }

    if ($wf.PSObject.Properties.Name -contains 'ModeScaleSelected') { $wf.ModeScaleSelected = $true }
    if ($wf.PSObject.Properties.Name -contains 'ModeWidthSelected') { $wf.ModeWidthSelected = $false }
    if ($wf.PSObject.Properties.Name -contains 'ModeHeightSelected') { $wf.ModeHeightSelected = $false }
    if ($wf.PSObject.Properties.Name -contains 'ModeFitToDisplaySelected') { $wf.ModeFitToDisplaySelected = $false }
    if ($wf.PSObject.Properties.Name -contains 'UpscaleScaleFactor') { $wf.UpscaleScaleFactor = [int]$ScaleFactor }
    else { $wf | Add-Member -Force -NotePropertyName UpscaleScaleFactor -NotePropertyValue ([int]$ScaleFactor) }
    if ($wf.PSObject.Properties.Name -contains 'LossyCompressionQuality') { $wf.LossyCompressionQuality = [int]$LossyQuality }
    else { $wf | Add-Member -Force -NotePropertyName LossyCompressionQuality -NotePropertyValue ([int]$LossyQuality) }
    if ($wf.PSObject.Properties.Name -contains 'OverwriteExistingFiles') { $wf.OverwriteExistingFiles = $true }
    else { $wf | Add-Member -Force -NotePropertyName OverwriteExistingFiles -NotePropertyValue $true }

    if ($wf.PSObject.Properties.Name -contains 'WorkflowOverrides') {
        if ($null -eq $wf.WorkflowOverrides) {
            $wf.WorkflowOverrides = [pscustomobject]@{}
        }
    }
    else {
        $wf | Add-Member -Force -NotePropertyName WorkflowOverrides -NotePropertyValue ([pscustomobject]@{})
    }
    if (Test-ObjectHasProperty -Object $wf.WorkflowOverrides -PropertyName 'GrayscaleDetectionThreshold') {
        $wf.WorkflowOverrides.GrayscaleDetectionThreshold = [int]$GrayThreshold
    }
    else {
        $wf.WorkflowOverrides | Add-Member -Force -NotePropertyName GrayscaleDetectionThreshold -NotePropertyValue ([int]$GrayThreshold)
    }

    $fmt = $OutputFormat.ToLowerInvariant()
    foreach ($name in @('WebpSelected', 'PngSelected', 'JpegSelected', 'AvifSelected')) {
        if (-not ($wf.PSObject.Properties.Name -contains $name)) {
            $wf | Add-Member -Force -NotePropertyName $name -NotePropertyValue $false
        }
    }
    $wf.WebpSelected = $fmt -eq 'webp'
    $wf.PngSelected = $fmt -eq 'png'
    $wf.JpegSelected = $fmt -eq 'jpeg'
    $wf.AvifSelected = $fmt -eq 'avif'

    Write-JsonUtf8NoBom -Object $cfg -Path $RuntimePath
}

function Invoke-BackendUpscaleGroup {
    param(
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$BackendScript,
        [Parameter(Mandatory = $true)][object]$BaseSettings,
        [Parameter(Mandatory = $true)][string]$ModelsDirectory,
        [Parameter(Mandatory = $true)][int]$ScaleFactor,
        [Parameter(Mandatory = $true)][int]$LossyQuality,
        [Parameter(Mandatory = $true)][int]$GrayThreshold,
        [Parameter(Mandatory = $true)][string]$OutputFormat,
        [Parameter(Mandatory = $true)][object[]]$Entries,
        [Parameter(Mandatory = $true)][string]$ExtractRoot,
        [Parameter(Mandatory = $true)][string]$WorkingRoot,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [scriptblock]$ProgressLineCallback
    )
    if ($Entries.Count -lt 1) {
        return [pscustomobject]@{ ExitCode = 0; RenameMap = @{} }
    }

    $groupToken = "{0}_{1}" -f $OutputFormat.ToLowerInvariant(), ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $inputRoot = Join-Path $WorkingRoot ("input_" + $groupToken)
    $outputRoot = Join-Path $WorkingRoot ("output_" + $groupToken)
    New-Item -ItemType Directory -Force -Path $inputRoot | Out-Null
    New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

    foreach ($entry in $Entries) {
        $src = [string]$entry.FullPath
        $rel = [string]$entry.RelPath
        $dst = Join-Path $inputRoot $rel
        $dstDir = Split-Path -Parent $dst
        if (-not [string]::IsNullOrWhiteSpace($dstDir)) { New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }
        Copy-Item -LiteralPath $src -Destination $dst -Force
    }

    $runtimePath = Join-Path $WorkingRoot ("runtime_" + $groupToken + '.json')
    New-RuntimeSettingsFile -BaseSettings $BaseSettings -RuntimePath $runtimePath -InputFolderPath $inputRoot -OutputFolderPath $outputRoot -ModelsDirectory $ModelsDirectory -ScaleFactor $ScaleFactor -LossyQuality $LossyQuality -GrayThreshold $GrayThreshold -OutputFormat $OutputFormat

    $lines = New-Object 'System.Collections.Generic.List[string]'
    & $PythonExe $BackendScript --settings $runtimePath 2>&1 | ForEach-Object {
        $line = [string]$_
        $lines.Add($line) | Out-Null
        [System.IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
        if ($null -ne $ProgressLineCallback) {
            try {
                & $ProgressLineCallback $line
            }
            catch {
                if (Test-DebugLog) {
                    Write-DebugLine ("Progress callback failed: {0}" -f $_.Exception.Message)
                }
            }
        }
        if (Test-DebugLog) { Write-Host $line }
    }
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "Backend failed for format '$OutputFormat' (exit=$exitCode)."
    }

    $renameMap = @{}
    foreach ($entry in $Entries) {
        $rel = [string]$entry.RelPath
        $sourceExt = [string]$entry.SourceExtension
        $produced = Get-ProducedImagePath -GroupOutputRoot $outputRoot -RelPath $rel -Format $OutputFormat
        if (-not $produced) {
            throw "Upscaled image not found for '$rel' in output format '$OutputFormat'."
        }

        $newRel = $rel
        if ($sourceExt -eq '.bmp') {
            $newRel = [System.IO.Path]::ChangeExtension($rel, '.png')
            $renameMap[$rel] = $newRel
        }

        $targetPath = Join-Path $ExtractRoot $newRel
        $targetDir = Split-Path -Parent $targetPath
        if (-not [string]::IsNullOrWhiteSpace($targetDir)) { New-Item -ItemType Directory -Force -Path $targetDir | Out-Null }
        Copy-Item -LiteralPath $produced -Destination $targetPath -Force

        if (-not [string]::Equals($rel, $newRel, [System.StringComparison]::OrdinalIgnoreCase)) {
            $oldPath = Join-Path $ExtractRoot $rel
            if (Test-Path -LiteralPath $oldPath -PathType Leaf) {
                Remove-Item -LiteralPath $oldPath -Force
            }
        }
    }

    return [pscustomobject]@{
        ExitCode = 0
        RenameMap = $renameMap
    }
}

function Write-EpubArchive {
    param(
        [Parameter(Mandatory = $true)][string]$SourceDirectory,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )
    if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $mimetypePath = Join-Path $SourceDirectory 'mimetype'
        $mimetypeEntry = $zip.CreateEntry('mimetype', [System.IO.Compression.CompressionLevel]::NoCompression)
        $mimetypeStream = $mimetypeEntry.Open()
        try {
            if (Test-Path -LiteralPath $mimetypePath -PathType Leaf) {
                $bytes = [System.IO.File]::ReadAllBytes($mimetypePath)
                $mimetypeStream.Write($bytes, 0, $bytes.Length)
            }
            else {
                $bytes = [System.Text.Encoding]::ASCII.GetBytes('application/epub+zip')
                $mimetypeStream.Write($bytes, 0, $bytes.Length)
            }
        }
        finally {
            $mimetypeStream.Dispose()
        }

        $files = @(Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File -ErrorAction SilentlyContinue | Sort-Object FullName)
        foreach ($file in $files) {
            $rel = Get-RelativePath -BasePath $SourceDirectory -TargetPath $file.FullName
            $relPosix = Convert-ToPosixPath -Path $rel
            if ([string]::Equals($relPosix, 'mimetype', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $entry = $zip.CreateEntry($relPosix, [System.IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $entry.Open()
            $srcStream = [System.IO.File]::OpenRead($file.FullName)
            try { $srcStream.CopyTo($entryStream) } finally { $srcStream.Dispose(); $entryStream.Dispose() }
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Invoke-EpubUpscaleFile {
    param(
        [Parameter(Mandatory = $true)][string]$InputEpubPath,
        [Parameter(Mandatory = $true)][string]$OutputEpubPath,
        [Parameter(Mandatory = $true)][object]$BaseSettings,
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$BackendScript,
        [Parameter(Mandatory = $true)][string]$ModelsDirectory,
        [Parameter(Mandatory = $true)][int]$ScaleFactor,
        [Parameter(Mandatory = $true)][int]$LossyQuality,
        [Parameter(Mandatory = $true)][int]$GrayThreshold,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [scriptblock]$ProgressCallback
    )
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("manga_epub_automation_epub_" + [guid]::NewGuid().ToString('N'))
    $extractRoot = Join-Path $tempRoot 'extract'
    $workingRoot = Join-Path $tempRoot 'work'
    New-Item -ItemType Directory -Force -Path $extractRoot | Out-Null
    New-Item -ItemType Directory -Force -Path $workingRoot | Out-Null

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($InputEpubPath, $extractRoot)

        $imageInfo = Get-EpubImageEntries -ExtractRoot $extractRoot
        $entries = Convert-ToObjectArray -Value $imageInfo.Entries
        if ($entries.Count -lt 1) {
            throw "No image entries found in EPUB: $InputEpubPath"
        }

        $fileProgressState = [pscustomobject]@{
            TotalUnits = [int]$entries.Count
            DoneUnits = 0
            ProcessedUnits = 0
            SkippedUnits = 0
        }
        $emitProgress = {
            param([string]$Status = 'processing')
            if ($null -eq $ProgressCallback) { return }
            & $ProgressCallback ([pscustomobject]@{
                    total_units = [int]$fileProgressState.TotalUnits
                    done_units = [int]$fileProgressState.DoneUnits
                    processed_units = [int]$fileProgressState.ProcessedUnits
                    skipped_units = [int]$fileProgressState.SkippedUnits
                    status = $Status
                })
        }
        $onBackendLine = {
            param([string]$Line)
            if ([string]::IsNullOrWhiteSpace($Line)) { return }

            $countDone = $false
            if ($Line -match '(?:^|\s)PROGRESS=') {
                $fileProgressState.ProcessedUnits = [int]$fileProgressState.ProcessedUnits + 1
                $countDone = $true
            }
            elseif ($Line -like '*file exists, skip:*') {
                $fileProgressState.SkippedUnits = [int]$fileProgressState.SkippedUnits + 1
                $countDone = $true
            }

            if ($countDone) {
                if ([int]$fileProgressState.DoneUnits -lt [int]$fileProgressState.TotalUnits) {
                    $fileProgressState.DoneUnits = [int]$fileProgressState.DoneUnits + 1
                }
                & $emitProgress 'processing'
            }
        }
        & $emitProgress 'start'

        $groups = @{}
        foreach ($entry in $entries) {
            $fmt = Get-UpscaleFormatForExtension -Extension ([string]$entry.SourceExtension)
            if (-not $groups.ContainsKey($fmt)) { $groups[$fmt] = New-Object 'System.Collections.Generic.List[object]' }
            $groups[$fmt].Add($entry) | Out-Null
        }

        $renameMap = @{}
        foreach ($fmt in @($groups.Keys | Sort-Object)) {
            Write-DebugLine ("Upscaling format group '{0}' ({1} images)" -f $fmt, $groups[$fmt].Count)
            $groupEntries = Convert-ToObjectArray -Value $groups[$fmt]
            $groupResult = Invoke-BackendUpscaleGroup -PythonExe $PythonExe -BackendScript $BackendScript -BaseSettings $BaseSettings -ModelsDirectory $ModelsDirectory -ScaleFactor $ScaleFactor -LossyQuality $LossyQuality -GrayThreshold $GrayThreshold -OutputFormat $fmt -Entries $groupEntries -ExtractRoot $extractRoot -WorkingRoot $workingRoot -LogPath $LogPath -ProgressLineCallback $onBackendLine
            foreach ($k in @($groupResult.RenameMap.Keys)) {
                $renameMap[$k] = $groupResult.RenameMap[$k]
            }
        }

        if ($renameMap.Count -gt 0) {
            if ($imageInfo.OpfPath) {
                Update-XmlManifestReferences -OpfPath $imageInfo.OpfPath -ExtractRoot $extractRoot -RenameMap $renameMap
            }
            Update-TextReferenceFiles -ExtractRoot $extractRoot -RenameMap $renameMap
        }

        $outDir = Split-Path -Parent $OutputEpubPath
        if (-not [string]::IsNullOrWhiteSpace($outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
        Write-EpubArchive -SourceDirectory $extractRoot -OutputPath $OutputEpubPath

        if ([int]$fileProgressState.DoneUnits -lt [int]$fileProgressState.TotalUnits) {
            $fileProgressState.DoneUnits = [int]$fileProgressState.TotalUnits
        }
        & $emitProgress 'completed'
        return [pscustomobject]@{
            TotalUnits = [int]$fileProgressState.TotalUnits
            DoneUnits = [int]$fileProgressState.DoneUnits
            ProcessedUnits = [int]$fileProgressState.ProcessedUnits
            SkippedUnits = [int]$fileProgressState.SkippedUnits
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot -PathType Container) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Update-LatestArtifact {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$LatestPath
    )
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return }
    Copy-Item -LiteralPath $SourcePath -Destination $LatestPath -Force
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$logsDir = Join-Path $scriptRoot 'logs'
New-Item -ItemType Directory -Force -Path $logsDir | Out-Null

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $logsDir ("epub_upscale_run_{0}.log" -f $timestamp)
$planPath = Join-Path $logsDir ("epub_run_plan_{0}.json" -f $timestamp)
$resultPath = Join-Path $logsDir ("epub_run_result_{0}.json" -f $timestamp)
$latestPlanPath = Join-Path $logsDir 'latest_epub_run_plan.json'
$latestResultPath = Join-Path $logsDir 'latest_epub_run_result.json'

Write-Info ("Log: {0}" -f $logPath)

$depsResolved = Resolve-DependenciesConfig -ScriptRoot $scriptRoot -DepsConfigPath $DepsConfigPath
$pipelineConfig = Resolve-PipelineConfig -ScriptRoot $scriptRoot -ConfigPath $ConfigPath

$effectiveScale = if ($PSBoundParameters.ContainsKey('UpscaleFactor')) { [int]$UpscaleFactor } else { [int]$pipelineConfig.Manga.UpscaleScaleFactor }
$effectiveQuality = if ($PSBoundParameters.ContainsKey('LossyQuality')) { [int]$LossyQuality } else { [int]$pipelineConfig.Manga.LossyCompressionQuality }
$effectiveGrayThreshold = if ($PSBoundParameters.ContainsKey('GrayscaleDetectionThreshold')) { [int]$GrayscaleDetectionThreshold } else { [int]$pipelineConfig.Manga.GrayscaleDetectionThreshold }
$effectiveOutputSuffix = if ($NoOutputSuffix -or [string]::IsNullOrWhiteSpace($OutputSuffix)) { '' } else { [string]$OutputSuffix }

$inputMode = if ($PSCmdlet.ParameterSetName -eq 'ByFiles') { 'files' } else { 'directory' }
$inputFiles = @(Get-InputEpubFiles -ParameterSetName $PSCmdlet.ParameterSetName -InputEpubPath $InputEpubPath -InputEpubDirectory $InputEpubDirectory)
$resolvedOutputDirectory = ''
$outputDirectoryError = $null
try {
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $resolvedOutputDirectory = [System.IO.Path]::GetFullPath((Get-Location).Path)
    }
    else {
        $resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
    }
}
catch {
    $outputDirectoryError = $_.Exception.Message
}

$planItems = New-Object 'System.Collections.Generic.List[object]'
foreach ($input in $inputFiles) {
    $exists = Test-Path -LiteralPath $input -PathType Leaf
    $isEpub = $exists -and ([string]::Equals([System.IO.Path]::GetExtension($input), '.epub', [System.StringComparison]::OrdinalIgnoreCase))
    $baseName = ''
    try {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension([string]$input)
    }
    catch {
        $baseName = ''
    }
    $outputPath = if (-not [string]::IsNullOrWhiteSpace($resolvedOutputDirectory) -and -not [string]::IsNullOrWhiteSpace($baseName)) {
        Join-Path $resolvedOutputDirectory ($baseName + $effectiveOutputSuffix + '.epub')
    }
    else { '' }
    $willSkip = $exists -and $isEpub -and (Test-Path -LiteralPath $outputPath -PathType Leaf) -and (-not $OverwriteOutputEpub)
    $planItems.Add([pscustomobject]@{
            input_path = $input
            input_exists = [bool]$exists
            input_is_epub = [bool]$isEpub
            output_path = $outputPath
            output_exists = if ([string]::IsNullOrWhiteSpace($outputPath)) { $false } else { Test-Path -LiteralPath $outputPath -PathType Leaf }
            predicted_skip = [bool]$willSkip
        }) | Out-Null
}

$preflightIssues = New-Object 'System.Collections.Generic.List[object]'
if ($inputFiles.Count -lt 1) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'INPUT_EMPTY' -Message 'No EPUB inputs resolved from current mode.')) | Out-Null
}
if ($null -ne $outputDirectoryError) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'OUTPUT_DIR_INVALID' -Message ("Output directory invalid: {0}" -f $outputDirectoryError))) | Out-Null
}
elseif ([string]::IsNullOrWhiteSpace($resolvedOutputDirectory)) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'OUTPUT_DIR_EMPTY' -Message 'Output directory resolved to empty value.')) | Out-Null
}
elseif (Test-Path -LiteralPath $resolvedOutputDirectory -PathType Leaf) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'OUTPUT_DIR_IS_FILE' -Message ("Output directory points to a file: {0}" -f $resolvedOutputDirectory))) | Out-Null
}
foreach ($item in $planItems) {
    if (-not $item.input_exists) {
        $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'INPUT_NOT_FOUND' -Message ("Input not found: {0}" -f $item.input_path))) | Out-Null
        continue
    }
    if (-not $item.input_is_epub) {
        $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'INPUT_NOT_EPUB' -Message ("Input is not .epub: {0}" -f $item.input_path))) | Out-Null
    }
}
$duplicateOutputGroups = @(
    $planItems |
    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.output_path) } |
    Group-Object output_path |
    Where-Object { $_.Count -gt 1 }
)
foreach ($dup in $duplicateOutputGroups) {
    $inputs = @($dup.Group | ForEach-Object { [string]$_.input_path })
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'OUTPUT_PATH_COLLISION' -Message ("Multiple inputs resolve to same output '{0}': {1}" -f [string]$dup.Name, ($inputs -join '; ')))) | Out-Null
}

if (-not $depsResolved.ConfigExists) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'DEPS_CONFIG_MISSING' -Message ("Deps config not found: {0}" -f $depsResolved.ConfigPath))) | Out-Null
}
elseif (-not $depsResolved.ConfigValid) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'DEPS_CONFIG_INVALID_JSON' -Message ("Deps config JSON invalid: {0}" -f $depsResolved.ConfigError))) | Out-Null
}

$pythonExePath = [string]$depsResolved.ResolvedPaths.python_exe
$backendScriptPath = [string]$depsResolved.ResolvedPaths.backend_script
$modelsDirPath = [string]$depsResolved.ResolvedPaths.models_dir
$baseSettingsPath = Join-Path $scriptRoot 'manga_epub_automation.settings.json'

if ([string]::IsNullOrWhiteSpace($pythonExePath) -or -not (Test-Path -LiteralPath $pythonExePath -PathType Leaf)) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'DEPS_PATH_PYTHON' -Message ("python_exe missing or invalid: {0}" -f $pythonExePath))) | Out-Null
}
if ([string]::IsNullOrWhiteSpace($backendScriptPath) -or -not (Test-Path -LiteralPath $backendScriptPath -PathType Leaf)) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'DEPS_PATH_BACKEND' -Message ("backend_script missing or invalid: {0}" -f $backendScriptPath))) | Out-Null
}
if ([string]::IsNullOrWhiteSpace($modelsDirPath) -or -not (Test-Path -LiteralPath $modelsDirPath -PathType Container)) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'DEPS_PATH_MODELS' -Message ("models_dir missing or invalid: {0}" -f $modelsDirPath))) | Out-Null
}
if (-not (Test-Path -LiteralPath $baseSettingsPath -PathType Leaf)) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'ERROR' -Code 'SETTINGS_TEMPLATE_MISSING' -Message ("Settings template not found: {0}" -f $baseSettingsPath))) | Out-Null
}

$predictedSkip = @($planItems | Where-Object { $_.predicted_skip }).Count
$predictedProcess = @($planItems | Where-Object { $_.input_exists -and $_.input_is_epub -and (-not $_.predicted_skip) }).Count
if ($predictedProcess -eq 0 -and $planItems.Count -gt 0) {
    $preflightIssues.Add((New-PreflightIssue -Severity 'INFO' -Code 'ALL_SKIP_PREDICTED' -Message 'All resolved EPUB files are predicted to be skipped with current overwrite setting.')) | Out-Null
}

if ($FailOnPreflightWarnings) {
    foreach ($issue in @($preflightIssues)) {
        if ([string]::Equals([string]$issue.Severity, 'WARN', [System.StringComparison]::OrdinalIgnoreCase)) {
            $issue.Severity = 'ERROR'
            $issue.Code = 'WARN_ESCALATED_' + [string]$issue.Code
        }
    }
}

$errorCount = @($preflightIssues | Where-Object { $_.Severity -eq 'ERROR' }).Count
$warnCount = @($preflightIssues | Where-Object { $_.Severity -eq 'WARN' }).Count
$infoCount = @($preflightIssues | Where-Object { $_.Severity -eq 'INFO' }).Count

$inputFilesArray = [object[]]@($inputFiles)
$inputPreview = [object[]]($inputFilesArray | Select-Object -First 20)
$planItemsArray = [object[]]$planItems.ToArray()
$preflightIssuesArray = [object[]]$preflightIssues.ToArray()

$planObject = [pscustomobject]@{
    TimestampUtc = [DateTime]::UtcNow.ToString('o')
    InputMode = $inputMode
    InputCount = [int]$inputFilesArray.Count
    InputPreview = $inputPreview
    InputFiles = $inputFilesArray
    OutputDirectory = $resolvedOutputDirectory
    OutputDirectoryInput = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { '' } else { $OutputDirectory }
    OutputSuffix = $effectiveOutputSuffix
    OverwriteOutputEpub = [bool]$OverwriteOutputEpub
    EffectiveSettings = [pscustomobject]@{
        UpscaleFactor = [int]$effectiveScale
        LossyQuality = [int]$effectiveQuality
        GrayscaleDetectionThreshold = [int]$effectiveGrayThreshold
    }
    Dependencies = [pscustomobject]@{
        DepsConfigPath = $depsResolved.ConfigPath
        DepsConfigExists = [bool]$depsResolved.ConfigExists
        DepsConfigValid = [bool]$depsResolved.ConfigValid
        PythonExe = $pythonExePath
        BackendScript = $backendScriptPath
        ModelsDir = $modelsDirPath
    }
    Forecast = [pscustomobject]@{
        planned = [int]$planItems.Count
        predicted_process = [int]$predictedProcess
        predicted_skip = [int]$predictedSkip
    }
    Items = $planItemsArray
    Preflight = [pscustomobject]@{
        ErrorCount = [int]$errorCount
        WarnCount = [int]$warnCount
        InfoCount = [int]$infoCount
        Issues = $preflightIssuesArray
    }
}

Write-JsonUtf8NoBom -Object $planObject -Path $planPath
Update-LatestArtifact -SourcePath $planPath -LatestPath $latestPlanPath
Write-GuiEvent -Type 'plan_ready' -Data ([pscustomobject]@{ plan_path = $planPath })
Write-GuiEvent -Type 'preflight_summary' -Data ([pscustomobject]@{
        errors = [int]$errorCount
        warnings = [int]$warnCount
        infos = [int]$infoCount
        issues = $preflightIssuesArray
    })

Write-Host ''
Write-Host '========== EPUB Upscale Plan =========='
Write-Host ("Mode:            {0}" -f $inputMode)
Write-Host ("InputCount:      {0}" -f $planItems.Count)
Write-Host ("Process/Skip:    {0}/{1}" -f $predictedProcess, $predictedSkip)
Write-Host ("OutputDir:       {0}" -f $resolvedOutputDirectory)
$outputSuffixDisplay = if ([string]::IsNullOrEmpty($effectiveOutputSuffix)) { '<empty>' } else { $effectiveOutputSuffix }
Write-Host ("OutputSuffix:    {0}" -f $outputSuffixDisplay)
Write-Host ("OverwriteOutput: {0}" -f ([bool]$OverwriteOutputEpub))
Write-Host ("Scale/Quality:   {0}/{1}" -f $effectiveScale, $effectiveQuality)
Write-Host ("GrayThreshold:   {0}" -f $effectiveGrayThreshold)
Write-Host ("Plan file:       {0}" -f $planPath)
Write-Host '======================================='

if ($errorCount -gt 0) {
    foreach ($issue in @($preflightIssues | Where-Object { $_.Severity -eq 'ERROR' })) {
        Write-Warn ("[{0}] {1}" -f $issue.Code, $issue.Message)
    }
    $result = [pscustomobject]@{
        TimestampUtc = [DateTime]::UtcNow.ToString('o')
        Status = 'preflight_failed'
        Message = 'Preflight contains blocking errors.'
        ProcessedFiles = 0
        SkippedFiles = 0
        FailedFiles = 0
        FailedFileList = @()
        OutputDirectory = $resolvedOutputDirectory
        PlanPath = $planPath
        LogPath = $logPath
    }
    Write-JsonUtf8NoBom -Object $result -Path $resultPath
    Update-LatestArtifact -SourcePath $resultPath -LatestPath $latestResultPath
    Write-GuiEvent -Type 'run_result' -Data ([pscustomobject]@{ result_path = $resultPath; status = 'preflight_failed' })
    throw 'Preflight failed. Fix errors and re-run.'
}

if ($PlanOnly) {
    $result = [pscustomobject]@{
        TimestampUtc = [DateTime]::UtcNow.ToString('o')
        Status = 'plan_only'
        Message = 'Plan-only mode completed.'
        ProcessedFiles = 0
        SkippedFiles = 0
        FailedFiles = 0
        FailedFileList = @()
        OutputDirectory = $resolvedOutputDirectory
        PlanPath = $planPath
        LogPath = $logPath
    }
    Write-JsonUtf8NoBom -Object $result -Path $resultPath
    Update-LatestArtifact -SourcePath $resultPath -LatestPath $latestResultPath
    Write-GuiEvent -Type 'run_result' -Data ([pscustomobject]@{ result_path = $resultPath; status = 'plan_only' })
    Write-Host 'PlanOnly completed.'
    exit 0
}

if (-not $AutoConfirm) {
    if ($GuiMode) {
        throw 'Interactive confirmation required. Re-run GUI flow with -AutoConfirm.'
    }
    $confirm = Read-Host 'Continue? [y/N]'
    if ($confirm -notin @('y', 'Y', 'yes', 'YES')) {
        $result = [pscustomobject]@{
            TimestampUtc = [DateTime]::UtcNow.ToString('o')
            Status = 'canceled'
            Message = 'Canceled by user.'
            ProcessedFiles = 0
            SkippedFiles = 0
            FailedFiles = 0
            FailedFileList = @()
            OutputDirectory = $resolvedOutputDirectory
            PlanPath = $planPath
            LogPath = $logPath
        }
        Write-JsonUtf8NoBom -Object $result -Path $resultPath
        Update-LatestArtifact -SourcePath $resultPath -LatestPath $latestResultPath
        Write-GuiEvent -Type 'run_result' -Data ([pscustomobject]@{ result_path = $resultPath; status = 'canceled' })
        Write-Host 'Canceled.'
        exit 0
    }
}

$baseSettings = Read-JsonFile -Path $baseSettingsPath
$totalFiles = [int]$planItems.Count
$processedFiles = 0
$skippedFiles = 0
$failedFiles = 0
$failedFileList = New-Object 'System.Collections.Generic.List[object]'
$doneFiles = 0
$interactive = [Environment]::UserInteractive
$currentFileTotalUnits = 0
$currentFileDoneUnits = 0
$currentFileProcessedUnits = 0
$currentFileSkippedUnits = 0

foreach ($item in $planItems) {
    $inputPath = [string]$item.input_path
    $outputPath = [string]$item.output_path

    $currentFileTotalUnits = 0
    $currentFileDoneUnits = 0
    $currentFileProcessedUnits = 0
    $currentFileSkippedUnits = 0

    $currentStatus = 'processing'
    $fileProgressState = [pscustomobject]@{
        TotalUnits = 0
        DoneUnits = 0
        ProcessedUnits = 0
        SkippedUnits = 0
    }
    $emitUiProgress = {
        param([string]$StatusForEvent)
        $currentFileTotalUnits = [int]$fileProgressState.TotalUnits
        $currentFileDoneUnits = [int]$fileProgressState.DoneUnits
        $currentFileProcessedUnits = [int]$fileProgressState.ProcessedUnits
        $currentFileSkippedUnits = [int]$fileProgressState.SkippedUnits
        $fraction = if ($currentFileTotalUnits -gt 0) { [double]$currentFileDoneUnits / [double]$currentFileTotalUnits } else { 0.0 }
        if ($fraction -lt 0.0) { $fraction = 0.0 }
        if ($fraction -gt 1.0) { $fraction = 1.0 }
        $percentDetailed = [int][math]::Round((($doneFiles + $fraction) * 100.0) / [math]::Max(1, $totalFiles))
        $statusLine = ("files {0}/{1} | current_images={2}/{3} | processed={4} skipped={5} failed={6}" -f $doneFiles, $totalFiles, $currentFileDoneUnits, $currentFileTotalUnits, $processedFiles, $skippedFiles, $failedFiles)
        if ($interactive) {
            Write-Progress -Id 31 -Activity 'EPUB Upscale Pipeline' -Status $statusLine -PercentComplete $percentDetailed
        }
        Write-GuiEvent -Type 'epub_file_progress' -Data ([pscustomobject]@{
                done_files = [int]$doneFiles
                total_files = [int]$totalFiles
                current_file = $inputPath
                percent = [int]$percentDetailed
                status = $StatusForEvent
                file_done_units = [int]$currentFileDoneUnits
                file_total_units = [int]$currentFileTotalUnits
                file_processed_units = [int]$currentFileProcessedUnits
                file_skipped_units = [int]$currentFileSkippedUnits
            })
    }

    try {
        if ($item.predicted_skip) {
            $skippedFiles += 1
            $currentStatus = 'skipped'
            Write-Info ("EPUB_SKIP: {0}" -f $inputPath)
        }
        elseif ($DryRun) {
            $currentStatus = 'dryrun'
            Write-Info ("EPUB_DRYRUN: {0}" -f $inputPath)
        }
        else {
            Write-Info ("EPUB_PROCESS: {0}" -f $inputPath)
            $fileProgressCallback = {
                param([object]$ProgressData)
                if ($null -eq $ProgressData) { return }
                if ($ProgressData.PSObject.Properties.Name -contains 'total_units') { $fileProgressState.TotalUnits = [int]$ProgressData.total_units }
                if ($ProgressData.PSObject.Properties.Name -contains 'done_units') { $fileProgressState.DoneUnits = [int]$ProgressData.done_units }
                if ($ProgressData.PSObject.Properties.Name -contains 'processed_units') { $fileProgressState.ProcessedUnits = [int]$ProgressData.processed_units }
                if ($ProgressData.PSObject.Properties.Name -contains 'skipped_units') { $fileProgressState.SkippedUnits = [int]$ProgressData.skipped_units }
                & $emitUiProgress 'processing'
            }
            $fileResult = Invoke-EpubUpscaleFile -InputEpubPath $inputPath -OutputEpubPath $outputPath -BaseSettings $baseSettings -PythonExe $pythonExePath -BackendScript $backendScriptPath -ModelsDirectory $modelsDirPath -ScaleFactor $effectiveScale -LossyQuality $effectiveQuality -GrayThreshold $effectiveGrayThreshold -LogPath $logPath -ProgressCallback $fileProgressCallback
            if ($null -ne $fileResult) {
                if ($fileResult.PSObject.Properties.Name -contains 'TotalUnits') { $fileProgressState.TotalUnits = [int]$fileResult.TotalUnits }
                if ($fileResult.PSObject.Properties.Name -contains 'DoneUnits') { $fileProgressState.DoneUnits = [int]$fileResult.DoneUnits }
                if ($fileResult.PSObject.Properties.Name -contains 'ProcessedUnits') { $fileProgressState.ProcessedUnits = [int]$fileResult.ProcessedUnits }
                if ($fileResult.PSObject.Properties.Name -contains 'SkippedUnits') { $fileProgressState.SkippedUnits = [int]$fileResult.SkippedUnits }
            }
            $processedFiles += 1
            $currentStatus = 'processed'
        }
    }
    catch {
        $failedFiles += 1
        $currentStatus = 'failed'
        $message = $_.Exception.Message
        $stack = $_.ScriptStackTrace
        $failedFileList.Add([pscustomobject]@{
                input_path = $inputPath
                output_path = $outputPath
                error = $message
            }) | Out-Null
        Write-Warn ("EPUB_FAIL: {0} ; {1}" -f $inputPath, $message)
        [System.IO.File]::AppendAllText($logPath, ("EPUB_FAIL: {0} ; {1}{2}" -f $inputPath, $message, [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
        if (-not [string]::IsNullOrWhiteSpace($stack)) {
            [System.IO.File]::AppendAllText($logPath, ("EPUB_FAIL_STACK: {0}{1}" -f $stack, [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
            if (Test-DebugLog) {
                Write-DebugLine ("EPUB_FAIL_STACK: {0}" -f $stack)
            }
        }
    }

    $doneFiles += 1
    $currentFileTotalUnits = [int]$fileProgressState.TotalUnits
    $currentFileDoneUnits = [int]$fileProgressState.DoneUnits
    $currentFileProcessedUnits = [int]$fileProgressState.ProcessedUnits
    $currentFileSkippedUnits = [int]$fileProgressState.SkippedUnits
    $percent = [int][math]::Round(($doneFiles * 100.0) / [math]::Max(1, $totalFiles))
    $statusLine = ("files {0}/{1} | processed={2} skipped={3} failed={4}" -f $doneFiles, $totalFiles, $processedFiles, $skippedFiles, $failedFiles)
    if ($currentStatus -eq 'processed' -and $currentFileTotalUnits -gt 0 -and $currentFileDoneUnits -lt $currentFileTotalUnits) {
        $currentFileDoneUnits = $currentFileTotalUnits
        $fileProgressState.DoneUnits = $currentFileDoneUnits
    }
    if ($interactive) {
        Write-Progress -Id 31 -Activity 'EPUB Upscale Pipeline' -Status $statusLine -PercentComplete $percent
    }
    else {
        Write-Host ("EPUB_FILE_PROGRESS: {0}% {1}" -f $percent, $statusLine)
    }

    Write-GuiEvent -Type 'epub_file_progress' -Data ([pscustomobject]@{
            done_files = [int]$doneFiles
            total_files = [int]$totalFiles
            current_file = $inputPath
            percent = [int]$percent
            status = $currentStatus
            file_done_units = [int]$currentFileDoneUnits
            file_total_units = [int]$currentFileTotalUnits
            file_processed_units = [int]$currentFileProcessedUnits
            file_skipped_units = [int]$currentFileSkippedUnits
        })
}

if ($interactive) {
    Write-Progress -Id 31 -Activity 'EPUB Upscale Pipeline' -Completed
}

$status = if ($failedFiles -gt 0) { 'completed_with_failures' } else { 'completed' }
$failedFileListArray = [object[]]$failedFileList.ToArray()
$resultObject = [pscustomobject]@{
    TimestampUtc = [DateTime]::UtcNow.ToString('o')
    Status = $status
    Message = if ($failedFiles -gt 0) { "Completed with failures: $failedFiles file(s)." } else { 'Completed successfully.' }
    ProcessedFiles = [int]$processedFiles
    SkippedFiles = [int]$skippedFiles
    FailedFiles = [int]$failedFiles
    FailedFileList = $failedFileListArray
    InputMode = $inputMode
    InputCount = [int]$totalFiles
    OutputDirectory = $resolvedOutputDirectory
    OutputSuffix = $effectiveOutputSuffix
    OverwriteOutputEpub = [bool]$OverwriteOutputEpub
    EffectiveSettings = [pscustomobject]@{
        UpscaleFactor = [int]$effectiveScale
        LossyQuality = [int]$effectiveQuality
        GrayscaleDetectionThreshold = [int]$effectiveGrayThreshold
    }
    PlanPath = $planPath
    LogPath = $logPath
}

Write-JsonUtf8NoBom -Object $resultObject -Path $resultPath
Update-LatestArtifact -SourcePath $resultPath -LatestPath $latestResultPath
Write-GuiEvent -Type 'run_result' -Data ([pscustomobject]@{
        result_path = $resultPath
        status = $status
        processed_files = [int]$processedFiles
        skipped_files = [int]$skippedFiles
        failed_files = [int]$failedFiles
    })

Write-Host ''
Write-Host '========== EPUB Upscale Result =========='
Write-Host ("Processed:      {0}" -f $processedFiles)
Write-Host ("Skipped:        {0}" -f $skippedFiles)
Write-Host ("Failed:         {0}" -f $failedFiles)
Write-Host ("Result file:    {0}" -f $resultPath)
Write-Host ('=========================================')

if ($failedFiles -gt 0) {
    throw "EPUB upscale pipeline completed with $failedFiles failure(s). See log: $logPath"
}
