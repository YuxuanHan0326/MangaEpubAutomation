<#
.SYNOPSIS
Automate MangaJaNai upscaling, chapter EPUB packaging, and merged EPUB generation.
#>
[CmdletBinding()]
param(
    [string]$TitleRoot,
    [string]$SourceDirName,
    [ValidateSet(1, 2, 3, 4)]
    [int]$UpscaleFactor = 2,
    [ValidateSet('webp', 'png', 'jpeg', 'avif')]
    [string]$OutputFormat = 'webp',
    [ValidateRange(1, 100)]
    [int]$LossyQuality = 80,
    [switch]$DryRun,
    [switch]$SkipUpscale,
    [switch]$SkipEpubPackaging,
    [switch]$SkipMergedEpub,
    [switch]$AutoConfirm,
    [switch]$PlanOnly,
    [switch]$MergePreviewCompact,
    [string]$MergeOrderFilePath = '',
    [switch]$DumpMergeOrderTemplate,
    [switch]$FailOnPreflightWarnings,
    [string]$KccExePath = '',
    [string]$DepsConfigPath = '',
    [string]$ConfigPath = '',
    [switch]$InitConfig,
    [switch]$InitDepsConfig,
    [switch]$NoUpscaleProgress,
    [switch]$GuiMode,
    [ValidateSet('info', 'debug')]
    [string]$LogLevel = 'info',
    [string]$EpubAuthorFallback = 'KCC',
    [string]$MergedEpubAuthorFallback = '',
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:PipelineGuiMode = [bool]$GuiMode
$script:LatestRunPlanPath = ''
$script:LatestRunResultPath = ''

$GroupTankobon = ([char]0x5355).ToString() + ([char]0x884C) + ([char]0x672C) # 单行本
$GroupDefault = ([char]0x9ED8).ToString() + ([char]0x8A8D) # 默認
$MetadataJsonName = ([char]0x5143).ToString() + ([char]0x6570) + ([char]0x636E) + '.json' # 元数据.json
$ChapterMetadataJsonName = ([char]0x7AE0).ToString() + ([char]0x8282) + ([char]0x5143) + ([char]0x6570) + ([char]0x636E) + '.json' # 章节元数据.json
$VolumeChar = [string][char]0x5377 # 卷
$TalkChar = [string][char]0x8BDD # 话

function Write-Info([string]$Message) { Write-Host "[INFO] $Message" }
function Write-Warn([string]$Message) { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Test-DebugLog { return $LogLevel -eq 'debug' }
function Write-DebugLine([string]$Message) { if (Test-DebugLog) { Write-Host "[DEBUG] $Message" -ForegroundColor DarkGray } }

function Write-GuiEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [object]$Data = $null
    )

    if (-not $script:PipelineGuiMode) { return }

    $eventObj = [pscustomobject]@{
        type = $Type
        ts_utc = (Get-Date).ToUniversalTime().ToString('o')
        data = $Data
    }
    try {
        $json = $eventObj | ConvertTo-Json -Depth 100 -Compress
    }
    catch {
        $fallback = [pscustomobject]@{
            type = 'event_serialize_error'
            ts_utc = (Get-Date).ToUniversalTime().ToString('o')
            data = [pscustomobject]@{
                source_type = $Type
                message = $_.Exception.Message
            }
        }
        $json = $fallback | ConvertTo-Json -Depth 10 -Compress
    }
    Write-Host ("PIPELINE_EVENT: {0}" -f $json)
}

function Update-LatestArtifact {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$LatestPath
    )

    if (-not (Test-PathSafe -Path $SourcePath -PathType Leaf)) { return }
    try {
        Copy-Item -LiteralPath $SourcePath -Destination $LatestPath -Force
    }
    catch {
        Write-DebugLine ("latest artifact update failed: src={0} dst={1} err={2}" -f $SourcePath, $LatestPath, $_.Exception.Message)
    }
}

function Write-GuiStageEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$Phase,
        [object]$Data = $null
    )

    Write-GuiEvent -Type 'stage' -Data ([pscustomobject]@{
            stage = $Stage
            phase = $Phase
            data = $Data
        })
}

function Test-PathSafe {
    param(
        [string]$Path,
        [ValidateSet('Any', 'Leaf', 'Container')]
        [string]$PathType = 'Any'
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        switch ($PathType) {
            'Leaf' { return (Test-Path -LiteralPath $Path -PathType Leaf) }
            'Container' { return (Test-Path -LiteralPath $Path -PathType Container) }
            default { return (Test-Path -LiteralPath $Path) }
        }
    }
    catch {
        return $false
    }
}

function Add-LogEntry {
    param(
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$Message
    )
    Add-Content -LiteralPath $LogPath -Value $Message -Encoding UTF8
}

function Write-StageMessage {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [switch]$DebugOnly,
        [switch]$Warning
    )

    Add-LogEntry -LogPath $LogPath -Message $Message
    if ($DebugOnly -and -not (Test-DebugLog)) { return }
    if ($Warning) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message }
}

function Invoke-ExternalWithLogging {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string]$ErrorPattern = '(?i)(traceback|exception|\berror\b|failed)'
    )

    $output = & $Executable @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    foreach ($item in @($output)) {
        $line = [string]$item
        Add-LogEntry -LogPath $LogPath -Message $line
        if (Test-DebugLog) { Write-Host $line }
        elseif ($line -match $ErrorPattern) { Write-Host $line -ForegroundColor Yellow }
    }
    return [pscustomobject]@{
        ExitCode = [int]$exitCode
        Lines = @($output | ForEach-Object { [string]$_ })
    }
}

function Get-DefaultDepsConfigObject {
    param([Parameter(Mandatory = $true)][string]$ScriptRoot)

    $appData = [string]$env:APPDATA
    $localAppData = [string]$env:LOCALAPPDATA
    $pythonDefault = if ($appData) { Join-Path $appData 'MangaJaNaiConverterGui\python\python\python.exe' } else { '' }
    $modelsDefault = if ($appData) { Join-Path $appData 'MangaJaNaiConverterGui\models' } else { '' }
    $backendDefault = if ($localAppData) { Join-Path $localAppData 'MangaJaNaiConverterGui\current\backend\src\run_upscale.py' } else { '' }

    return [pscustomobject]@{
        version = 2
        paths = [pscustomobject]@{
            python_exe = $pythonDefault
            backend_script = $backendDefault
            models_dir = $modelsDefault
            kcc_exe = Join-Path $ScriptRoot 'kcc_c2e_9.4.3.exe'
            merge_script = Join-Path $ScriptRoot 'MergeEpubByOrder.py'
        }
        progress = [pscustomobject]@{
            enabled = $true
            refresh_seconds = 1
            eta_min_samples = 8
            noninteractive_log_interval_seconds = 10
        }
    }
}

function Get-DefaultPipelineConfigObject {
    return [pscustomobject]@{
        version = 1
        pipeline = [pscustomobject]@{
            SourceDirSuffixToSkip = '-upscaled'
            OutputDirSuffix = '-upscaled'
            EpubOutputDirSuffix = '-output'
            OutputFilenamePattern = '%filename%-upscaled'
        }
        manga = [pscustomobject]@{
            SelectedTabIndex = 1
            OverwriteExistingFiles = $false
            ModeScaleSelected = $true
            ModeWidthSelected = $false
            ModeHeightSelected = $false
            ModeFitToDisplaySelected = $false
            UpscaleScaleFactor = 2
            OutputFormat = 'webp'
            LossyCompressionQuality = 80
            WorkflowOverrides = [pscustomobject]@{}
        }
        kcc = [pscustomobject]@{
            CliOptions = [pscustomobject]@{
                DeviceProfile = 'KS'
                OutputFormat = 'EPUB'
                NoKepub = $true
                DisableProcessing = $true
                ForceColor = $true
                MangaStyle = $false
                HighQualityMagnification = $false
                TwoPanelView = $false
                WebtoonMode = $false
                TargetSizeMB = $null
                LegacyPdfExtract = $false
                UpscaleSmallImages = $false
                StretchToResolution = $false
                SplitterMode = $null
                Gamma = $null
                CroppingMode = $null
                CroppingPower = $null
                PreserveMarginPercent = $null
                CroppingMinimumRatio = $null
                InterPanelCropMode = $null
                ForceBlackBorders = $false
                ForceWhiteBorders = $false
                ForcePng = $false
                MozJpeg = $false
                JpegQuality = $null
                MaximizeStrips = $false
                BatchSplitMode = $null
                MetadataTitleMode = $null
                SpreadShift = $false
                NoRotateSpreads = $false
                RotateFirstSpread = $false
                AutoLevel = $false
                DisableAutoContrast = $false
                ColorAutoContrast = $false
                FileFusion = $false
                EraseRainbow = $false
                DeleteSourceAfterPack = $false
                CustomWidth = $null
                CustomHeight = $null
                AdditionalArgs = @()
            }
            UnicodeStaging = [pscustomobject]@{
                Enabled = $true
                StagePrefix = 'manga_epub_automation_kcc_stage_'
                SafeTitleFallback = 'manga_epub_automation_book'
                SafeAuthorFallback = 'pipeline'
            }
            MetadataRewrite = [pscustomobject]@{
                Enabled = $true
            }
            BaseArgs = @()
        }
        merge = [pscustomobject]@{
            Language = 'zh'
            DescriptionHeader = '选集 包含:'
            IncludeOrderInDescription = $true
            MetadataContributor = 'manga-epub-automation-merge'
        }
    }
}

function Build-KccBaseArgsFromConfig {
    param([Parameter(Mandatory = $true)][object]$KccConfig)

    # Backward compatibility: if legacy BaseArgs is provided, keep precedence.
    if (($KccConfig.PSObject.Properties.Name -contains 'BaseArgs') -and $KccConfig.BaseArgs -and $KccConfig.BaseArgs.Count -gt 0) {
        return @($KccConfig.BaseArgs | ForEach-Object { [string]$_ })
    }

    $args = @()
    if (($KccConfig.PSObject.Properties.Name -contains 'CliOptions') -and $KccConfig.CliOptions) {
        $cli = $KccConfig.CliOptions

        $parseInt = {
            param($value)
            if ($null -eq $value) { return $null }
            $tmp = 0
            if ([int]::TryParse([string]$value, [ref]$tmp)) { return [int]$tmp }
            return $null
        }
        $parseDouble = {
            param($value)
            if ($null -eq $value) { return $null }
            $tmp = 0.0
            if ([double]::TryParse([string]$value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$tmp)) { return [double]$tmp }
            if ([double]::TryParse([string]$value, [ref]$tmp)) { return [double]$tmp }
            return $null
        }

        if (($cli.PSObject.Properties.Name -contains 'DeviceProfile') -and -not [string]::IsNullOrWhiteSpace([string]$cli.DeviceProfile)) {
            $args += @('-p', [string]$cli.DeviceProfile)
        }
        if (($cli.PSObject.Properties.Name -contains 'OutputFormat') -and -not [string]::IsNullOrWhiteSpace([string]$cli.OutputFormat)) {
            $args += @('-f', [string]$cli.OutputFormat)
        }
        if (($cli.PSObject.Properties.Name -contains 'MangaStyle') -and [bool]$cli.MangaStyle) { $args += '-m' }
        if (($cli.PSObject.Properties.Name -contains 'HighQualityMagnification') -and [bool]$cli.HighQualityMagnification) { $args += '-q' }
        if (($cli.PSObject.Properties.Name -contains 'TwoPanelView') -and [bool]$cli.TwoPanelView) { $args += '-2' }
        if (($cli.PSObject.Properties.Name -contains 'WebtoonMode') -and [bool]$cli.WebtoonMode) { $args += '-w' }

        if ($cli.PSObject.Properties.Name -contains 'TargetSizeMB') {
            $targetSize = & $parseInt $cli.TargetSizeMB
            if ($null -ne $targetSize -and $targetSize -gt 0) { $args += @('--ts', [string]$targetSize) }
        }

        if (($cli.PSObject.Properties.Name -contains 'NoKepub') -and [bool]$cli.NoKepub) {
            $args += '--nokepub'
        }
        if (($cli.PSObject.Properties.Name -contains 'DisableProcessing') -and [bool]$cli.DisableProcessing) {
            $args += '-n'
        }
        if (($cli.PSObject.Properties.Name -contains 'LegacyPdfExtract') -and [bool]$cli.LegacyPdfExtract) { $args += '--pdfextract' }
        if (($cli.PSObject.Properties.Name -contains 'UpscaleSmallImages') -and [bool]$cli.UpscaleSmallImages) { $args += '-u' }
        if (($cli.PSObject.Properties.Name -contains 'StretchToResolution') -and [bool]$cli.StretchToResolution) { $args += '-s' }

        if ($cli.PSObject.Properties.Name -contains 'SplitterMode') {
            $splitterMode = & $parseInt $cli.SplitterMode
            if ($null -ne $splitterMode -and $splitterMode -ge 0 -and $splitterMode -le 2) { $args += @('-r', [string]$splitterMode) }
        }
        if ($cli.PSObject.Properties.Name -contains 'Gamma') {
            $gammaText = [string]$cli.Gamma
            if (-not [string]::IsNullOrWhiteSpace($gammaText)) { $args += @('-g', $gammaText) }
        }
        if ($cli.PSObject.Properties.Name -contains 'CroppingMode') {
            $cropMode = & $parseInt $cli.CroppingMode
            if ($null -ne $cropMode -and $cropMode -ge 0 -and $cropMode -le 2) { $args += @('-c', [string]$cropMode) }
        }
        if ($cli.PSObject.Properties.Name -contains 'CroppingPower') {
            $cropPower = & $parseDouble $cli.CroppingPower
            if ($null -ne $cropPower) { $args += @('--cp', ([double]$cropPower).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture)) }
        }
        if ($cli.PSObject.Properties.Name -contains 'PreserveMarginPercent') {
            $preserveMargin = & $parseDouble $cli.PreserveMarginPercent
            if ($null -ne $preserveMargin) { $args += @('--preservemargin', ([double]$preserveMargin).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture)) }
        }
        if ($cli.PSObject.Properties.Name -contains 'CroppingMinimumRatio') {
            $cropMin = & $parseDouble $cli.CroppingMinimumRatio
            if ($null -ne $cropMin) { $args += @('--cm', ([double]$cropMin).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture)) }
        }
        if ($cli.PSObject.Properties.Name -contains 'InterPanelCropMode') {
            $ipcMode = & $parseInt $cli.InterPanelCropMode
            if ($null -ne $ipcMode -and $ipcMode -ge 0 -and $ipcMode -le 2) { $args += @('--ipc', [string]$ipcMode) }
        }
        if (($cli.PSObject.Properties.Name -contains 'ForceBlackBorders') -and [bool]$cli.ForceBlackBorders) { $args += '--blackborders' }
        if (($cli.PSObject.Properties.Name -contains 'ForceWhiteBorders') -and [bool]$cli.ForceWhiteBorders) { $args += '--whiteborders' }
        if (($cli.PSObject.Properties.Name -contains 'ForceColor') -and [bool]$cli.ForceColor) {
            $args += '--forcecolor'
        }
        if (($cli.PSObject.Properties.Name -contains 'ForcePng') -and [bool]$cli.ForcePng) { $args += '--forcepng' }
        if (($cli.PSObject.Properties.Name -contains 'MozJpeg') -and [bool]$cli.MozJpeg) { $args += '--mozjpeg' }
        if ($cli.PSObject.Properties.Name -contains 'JpegQuality') {
            $jpegQ = & $parseInt $cli.JpegQuality
            if ($null -ne $jpegQ -and $jpegQ -ge 0 -and $jpegQ -le 95) { $args += @('--jpeg-quality', [string]$jpegQ) }
        }
        if (($cli.PSObject.Properties.Name -contains 'MaximizeStrips') -and [bool]$cli.MaximizeStrips) { $args += '--maximizestrips' }
        if ($cli.PSObject.Properties.Name -contains 'BatchSplitMode') {
            $batchMode = & $parseInt $cli.BatchSplitMode
            if ($null -ne $batchMode -and $batchMode -ge 0 -and $batchMode -le 2) { $args += @('-b', [string]$batchMode) }
        }
        if ($cli.PSObject.Properties.Name -contains 'MetadataTitleMode') {
            $metaTitleMode = & $parseInt $cli.MetadataTitleMode
            if ($null -ne $metaTitleMode -and ($metaTitleMode -eq 1 -or $metaTitleMode -eq 2)) { $args += @('--metadatatitle', [string]$metaTitleMode) }
        }
        if (($cli.PSObject.Properties.Name -contains 'SpreadShift') -and [bool]$cli.SpreadShift) { $args += '--spreadshift' }
        if (($cli.PSObject.Properties.Name -contains 'NoRotateSpreads') -and [bool]$cli.NoRotateSpreads) { $args += '--norotate' }
        if (($cli.PSObject.Properties.Name -contains 'RotateFirstSpread') -and [bool]$cli.RotateFirstSpread) { $args += '--rotatefirst' }
        if (($cli.PSObject.Properties.Name -contains 'AutoLevel') -and [bool]$cli.AutoLevel) { $args += '--autolevel' }
        if (($cli.PSObject.Properties.Name -contains 'DisableAutoContrast') -and [bool]$cli.DisableAutoContrast) { $args += '--noautocontrast' }
        if (($cli.PSObject.Properties.Name -contains 'ColorAutoContrast') -and [bool]$cli.ColorAutoContrast) { $args += '--colorautocontrast' }
        if (($cli.PSObject.Properties.Name -contains 'FileFusion') -and [bool]$cli.FileFusion) { $args += '--filefusion' }
        if (($cli.PSObject.Properties.Name -contains 'EraseRainbow') -and [bool]$cli.EraseRainbow) { $args += '--eraserainbow' }
        if (($cli.PSObject.Properties.Name -contains 'DeleteSourceAfterPack') -and [bool]$cli.DeleteSourceAfterPack) { $args += '-d' }
        if ($cli.PSObject.Properties.Name -contains 'CustomWidth') {
            $cw = & $parseInt $cli.CustomWidth
            if ($null -ne $cw -and $cw -gt 0) { $args += @('--customwidth', [string]$cw) }
        }
        if ($cli.PSObject.Properties.Name -contains 'CustomHeight') {
            $ch = & $parseInt $cli.CustomHeight
            if ($null -ne $ch -and $ch -gt 0) { $args += @('--customheight', [string]$ch) }
        }
        if (($cli.PSObject.Properties.Name -contains 'AdditionalArgs') -and $cli.AdditionalArgs) {
            $args += @($cli.AdditionalArgs | ForEach-Object { [string]$_ })
        }
    }

    if ($args.Count -lt 1) {
        return @('-p', 'KS', '-f', 'EPUB', '--nokepub', '-n', '--forcecolor')
    }
    return @($args)
}

function Resolve-PipelineConfig {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [string]$ConfigPath
    )

    $defaults = Get-DefaultPipelineConfigObject
    $resolvedPath = if ([string]::IsNullOrWhiteSpace($ConfigPath)) { Join-Path $ScriptRoot 'manga_epub_automation.config.json' } else { $ConfigPath }
    $exists = Test-PathSafe -Path $resolvedPath -PathType Leaf
    $parsed = $null
    $valid = $false
    $errorText = $null

    if ($exists) {
        try {
            $parsed = Read-JsonFile -Path $resolvedPath
            $valid = $true
        }
        catch {
            $errorText = $_.Exception.Message
        }
    }

    if ($exists -and -not $valid) {
        throw "Invalid pipeline config JSON: $resolvedPath ; $errorText"
    }

    $cfg = [pscustomobject]@{
        version = 2
        pipeline = [pscustomobject]@{
            SourceDirSuffixToSkip = [string]$defaults.pipeline.SourceDirSuffixToSkip
            OutputDirSuffix = [string]$defaults.pipeline.OutputDirSuffix
            EpubOutputDirSuffix = [string]$defaults.pipeline.EpubOutputDirSuffix
            OutputFilenamePattern = [string]$defaults.pipeline.OutputFilenamePattern
        }
        manga = [pscustomobject]@{
            SelectedTabIndex = [int]$defaults.manga.SelectedTabIndex
            OverwriteExistingFiles = [bool]$defaults.manga.OverwriteExistingFiles
            ModeScaleSelected = [bool]$defaults.manga.ModeScaleSelected
            ModeWidthSelected = [bool]$defaults.manga.ModeWidthSelected
            ModeHeightSelected = [bool]$defaults.manga.ModeHeightSelected
            ModeFitToDisplaySelected = [bool]$defaults.manga.ModeFitToDisplaySelected
            UpscaleScaleFactor = [int]$defaults.manga.UpscaleScaleFactor
            OutputFormat = [string]$defaults.manga.OutputFormat
            LossyCompressionQuality = [int]$defaults.manga.LossyCompressionQuality
            WorkflowOverrides = [pscustomobject]@{}
        }
        kcc = [pscustomobject]@{
            CliOptions = [pscustomobject]@{
                DeviceProfile = [string]$defaults.kcc.CliOptions.DeviceProfile
                OutputFormat = [string]$defaults.kcc.CliOptions.OutputFormat
                NoKepub = [bool]$defaults.kcc.CliOptions.NoKepub
                DisableProcessing = [bool]$defaults.kcc.CliOptions.DisableProcessing
                ForceColor = [bool]$defaults.kcc.CliOptions.ForceColor
                MangaStyle = [bool]$defaults.kcc.CliOptions.MangaStyle
                HighQualityMagnification = [bool]$defaults.kcc.CliOptions.HighQualityMagnification
                TwoPanelView = [bool]$defaults.kcc.CliOptions.TwoPanelView
                WebtoonMode = [bool]$defaults.kcc.CliOptions.WebtoonMode
                TargetSizeMB = $defaults.kcc.CliOptions.TargetSizeMB
                LegacyPdfExtract = [bool]$defaults.kcc.CliOptions.LegacyPdfExtract
                UpscaleSmallImages = [bool]$defaults.kcc.CliOptions.UpscaleSmallImages
                StretchToResolution = [bool]$defaults.kcc.CliOptions.StretchToResolution
                SplitterMode = $defaults.kcc.CliOptions.SplitterMode
                Gamma = $defaults.kcc.CliOptions.Gamma
                CroppingMode = $defaults.kcc.CliOptions.CroppingMode
                CroppingPower = $defaults.kcc.CliOptions.CroppingPower
                PreserveMarginPercent = $defaults.kcc.CliOptions.PreserveMarginPercent
                CroppingMinimumRatio = $defaults.kcc.CliOptions.CroppingMinimumRatio
                InterPanelCropMode = $defaults.kcc.CliOptions.InterPanelCropMode
                ForceBlackBorders = [bool]$defaults.kcc.CliOptions.ForceBlackBorders
                ForceWhiteBorders = [bool]$defaults.kcc.CliOptions.ForceWhiteBorders
                ForcePng = [bool]$defaults.kcc.CliOptions.ForcePng
                MozJpeg = [bool]$defaults.kcc.CliOptions.MozJpeg
                JpegQuality = $defaults.kcc.CliOptions.JpegQuality
                MaximizeStrips = [bool]$defaults.kcc.CliOptions.MaximizeStrips
                BatchSplitMode = $defaults.kcc.CliOptions.BatchSplitMode
                MetadataTitleMode = $defaults.kcc.CliOptions.MetadataTitleMode
                SpreadShift = [bool]$defaults.kcc.CliOptions.SpreadShift
                NoRotateSpreads = [bool]$defaults.kcc.CliOptions.NoRotateSpreads
                RotateFirstSpread = [bool]$defaults.kcc.CliOptions.RotateFirstSpread
                AutoLevel = [bool]$defaults.kcc.CliOptions.AutoLevel
                DisableAutoContrast = [bool]$defaults.kcc.CliOptions.DisableAutoContrast
                ColorAutoContrast = [bool]$defaults.kcc.CliOptions.ColorAutoContrast
                FileFusion = [bool]$defaults.kcc.CliOptions.FileFusion
                EraseRainbow = [bool]$defaults.kcc.CliOptions.EraseRainbow
                DeleteSourceAfterPack = [bool]$defaults.kcc.CliOptions.DeleteSourceAfterPack
                CustomWidth = $defaults.kcc.CliOptions.CustomWidth
                CustomHeight = $defaults.kcc.CliOptions.CustomHeight
                AdditionalArgs = @($defaults.kcc.CliOptions.AdditionalArgs)
            }
            UnicodeStaging = [pscustomobject]@{
                Enabled = [bool]$defaults.kcc.UnicodeStaging.Enabled
                StagePrefix = [string]$defaults.kcc.UnicodeStaging.StagePrefix
                SafeTitleFallback = [string]$defaults.kcc.UnicodeStaging.SafeTitleFallback
                SafeAuthorFallback = [string]$defaults.kcc.UnicodeStaging.SafeAuthorFallback
            }
            MetadataRewrite = [pscustomobject]@{
                Enabled = [bool]$defaults.kcc.MetadataRewrite.Enabled
            }
            BaseArgs = @($defaults.kcc.BaseArgs)
            EffectiveBaseArgs = @($defaults.kcc.BaseArgs)
            EffectiveArgsSource = 'defaults'
        }
        merge = [pscustomobject]@{
            Language = [string]$defaults.merge.Language
            DescriptionHeader = [string]$defaults.merge.DescriptionHeader
            IncludeOrderInDescription = [bool]$defaults.merge.IncludeOrderInDescription
            MetadataContributor = [string]$defaults.merge.MetadataContributor
        }
    }

    if ($valid -and $parsed) {
        if (($parsed.PSObject.Properties.Name -contains 'pipeline') -and $parsed.pipeline) {
            $p = $parsed.pipeline
            if (($p.PSObject.Properties.Name -contains 'SourceDirSuffixToSkip') -and ($null -ne $p.SourceDirSuffixToSkip)) { $cfg.pipeline.SourceDirSuffixToSkip = [string]$p.SourceDirSuffixToSkip }
            if (($p.PSObject.Properties.Name -contains 'OutputDirSuffix') -and ($null -ne $p.OutputDirSuffix)) { $cfg.pipeline.OutputDirSuffix = [string]$p.OutputDirSuffix }
            if (($p.PSObject.Properties.Name -contains 'EpubOutputDirSuffix') -and ($null -ne $p.EpubOutputDirSuffix)) { $cfg.pipeline.EpubOutputDirSuffix = [string]$p.EpubOutputDirSuffix }
            if (($p.PSObject.Properties.Name -contains 'OutputFilenamePattern') -and $p.OutputFilenamePattern) { $cfg.pipeline.OutputFilenamePattern = [string]$p.OutputFilenamePattern }
        }

        if (($parsed.PSObject.Properties.Name -contains 'manga') -and $parsed.manga) {
            $m = $parsed.manga
            if (($m.PSObject.Properties.Name -contains 'SelectedTabIndex') -and $null -ne $m.SelectedTabIndex) { $cfg.manga.SelectedTabIndex = [int]$m.SelectedTabIndex }
            if ($m.PSObject.Properties.Name -contains 'OverwriteExistingFiles') { $cfg.manga.OverwriteExistingFiles = [bool]$m.OverwriteExistingFiles }
            if ($m.PSObject.Properties.Name -contains 'ModeScaleSelected') { $cfg.manga.ModeScaleSelected = [bool]$m.ModeScaleSelected }
            if ($m.PSObject.Properties.Name -contains 'ModeWidthSelected') { $cfg.manga.ModeWidthSelected = [bool]$m.ModeWidthSelected }
            if ($m.PSObject.Properties.Name -contains 'ModeHeightSelected') { $cfg.manga.ModeHeightSelected = [bool]$m.ModeHeightSelected }
            if ($m.PSObject.Properties.Name -contains 'ModeFitToDisplaySelected') { $cfg.manga.ModeFitToDisplaySelected = [bool]$m.ModeFitToDisplaySelected }
            if (($m.PSObject.Properties.Name -contains 'UpscaleScaleFactor') -and $null -ne $m.UpscaleScaleFactor) { $cfg.manga.UpscaleScaleFactor = [int]$m.UpscaleScaleFactor }
            if (($m.PSObject.Properties.Name -contains 'OutputFormat') -and $m.OutputFormat) { $cfg.manga.OutputFormat = ([string]$m.OutputFormat).ToLowerInvariant() }
            if (($m.PSObject.Properties.Name -contains 'LossyCompressionQuality') -and $null -ne $m.LossyCompressionQuality) { $cfg.manga.LossyCompressionQuality = [int]$m.LossyCompressionQuality }
            if (($m.PSObject.Properties.Name -contains 'WorkflowOverrides') -and $m.WorkflowOverrides) { $cfg.manga.WorkflowOverrides = $m.WorkflowOverrides }
        }

        if (($parsed.PSObject.Properties.Name -contains 'kcc') -and $parsed.kcc) {
            $k = $parsed.kcc
            if (($k.PSObject.Properties.Name -contains 'BaseArgs') -and $k.BaseArgs) { $cfg.kcc.BaseArgs = @($k.BaseArgs | ForEach-Object { [string]$_ }) }

            if (($k.PSObject.Properties.Name -contains 'CliOptions') -and $k.CliOptions) {
                $kc = $k.CliOptions
                foreach ($prop in $kc.PSObject.Properties) {
                    $pn = [string]$prop.Name
                    if ($cfg.kcc.CliOptions.PSObject.Properties.Name -notcontains $pn) { continue }
                    if ($pn -eq 'AdditionalArgs') {
                        if ($null -eq $prop.Value) { $cfg.kcc.CliOptions.AdditionalArgs = @() }
                        else { $cfg.kcc.CliOptions.AdditionalArgs = @($prop.Value | ForEach-Object { [string]$_ }) }
                    }
                    else {
                        $cfg.kcc.CliOptions.$pn = $prop.Value
                    }
                }
            }

            if (($k.PSObject.Properties.Name -contains 'UnicodeStaging') -and $k.UnicodeStaging) {
                $kus = $k.UnicodeStaging
                if ($kus.PSObject.Properties.Name -contains 'Enabled') { $cfg.kcc.UnicodeStaging.Enabled = [bool]$kus.Enabled }
                if (($kus.PSObject.Properties.Name -contains 'StagePrefix') -and $kus.StagePrefix) { $cfg.kcc.UnicodeStaging.StagePrefix = [string]$kus.StagePrefix }
                if (($kus.PSObject.Properties.Name -contains 'SafeTitleFallback') -and $kus.SafeTitleFallback) { $cfg.kcc.UnicodeStaging.SafeTitleFallback = [string]$kus.SafeTitleFallback }
                if (($kus.PSObject.Properties.Name -contains 'SafeAuthorFallback') -and $kus.SafeAuthorFallback) { $cfg.kcc.UnicodeStaging.SafeAuthorFallback = [string]$kus.SafeAuthorFallback }
            }

            if (($k.PSObject.Properties.Name -contains 'MetadataRewrite') -and $k.MetadataRewrite) {
                $kmr = $k.MetadataRewrite
                if ($kmr.PSObject.Properties.Name -contains 'Enabled') { $cfg.kcc.MetadataRewrite.Enabled = [bool]$kmr.Enabled }
            }

            # Backward compatibility for previous simple fields.
            if (($k.PSObject.Properties.Name -contains 'UnicodeStagePrefix') -and $k.UnicodeStagePrefix) { $cfg.kcc.UnicodeStaging.StagePrefix = [string]$k.UnicodeStagePrefix }
            if (($k.PSObject.Properties.Name -contains 'SafeTitleFallback') -and $k.SafeTitleFallback) { $cfg.kcc.UnicodeStaging.SafeTitleFallback = [string]$k.SafeTitleFallback }
            if (($k.PSObject.Properties.Name -contains 'SafeAuthorFallback') -and $k.SafeAuthorFallback) { $cfg.kcc.UnicodeStaging.SafeAuthorFallback = [string]$k.SafeAuthorFallback }
        }

        if (($parsed.PSObject.Properties.Name -contains 'merge') -and $parsed.merge) {
            $mg = $parsed.merge
            if (($mg.PSObject.Properties.Name -contains 'Language') -and $mg.Language) { $cfg.merge.Language = [string]$mg.Language }
            if (($mg.PSObject.Properties.Name -contains 'DescriptionHeader') -and $mg.DescriptionHeader) { $cfg.merge.DescriptionHeader = [string]$mg.DescriptionHeader }
            if ($mg.PSObject.Properties.Name -contains 'IncludeOrderInDescription') { $cfg.merge.IncludeOrderInDescription = [bool]$mg.IncludeOrderInDescription }
            if (($mg.PSObject.Properties.Name -contains 'MetadataContributor') -and $mg.MetadataContributor) { $cfg.merge.MetadataContributor = [string]$mg.MetadataContributor }
        }
    }

    if ([string]::IsNullOrWhiteSpace($cfg.pipeline.OutputFilenamePattern)) { $cfg.pipeline.OutputFilenamePattern = [string]$defaults.pipeline.OutputFilenamePattern }
    if ([string]::IsNullOrWhiteSpace($cfg.pipeline.SourceDirSuffixToSkip)) { $cfg.pipeline.SourceDirSuffixToSkip = [string]$defaults.pipeline.SourceDirSuffixToSkip }
    if ([string]::IsNullOrWhiteSpace($cfg.pipeline.OutputDirSuffix)) { $cfg.pipeline.OutputDirSuffix = [string]$defaults.pipeline.OutputDirSuffix }
    if ([string]::IsNullOrWhiteSpace($cfg.pipeline.EpubOutputDirSuffix)) { $cfg.pipeline.EpubOutputDirSuffix = [string]$defaults.pipeline.EpubOutputDirSuffix }
    if ($null -eq $cfg.kcc.CliOptions.AdditionalArgs) { $cfg.kcc.CliOptions.AdditionalArgs = @() }
    if ([string]::IsNullOrWhiteSpace($cfg.kcc.UnicodeStaging.StagePrefix)) { $cfg.kcc.UnicodeStaging.StagePrefix = [string]$defaults.kcc.UnicodeStaging.StagePrefix }
    if ([string]::IsNullOrWhiteSpace($cfg.kcc.UnicodeStaging.SafeTitleFallback)) { $cfg.kcc.UnicodeStaging.SafeTitleFallback = [string]$defaults.kcc.UnicodeStaging.SafeTitleFallback }
    if ([string]::IsNullOrWhiteSpace($cfg.kcc.UnicodeStaging.SafeAuthorFallback)) { $cfg.kcc.UnicodeStaging.SafeAuthorFallback = [string]$defaults.kcc.UnicodeStaging.SafeAuthorFallback }
    if ($cfg.manga.UpscaleScaleFactor -lt 1 -or $cfg.manga.UpscaleScaleFactor -gt 4) { $cfg.manga.UpscaleScaleFactor = [int]$defaults.manga.UpscaleScaleFactor }
    if ($cfg.manga.LossyCompressionQuality -lt 1 -or $cfg.manga.LossyCompressionQuality -gt 100) { $cfg.manga.LossyCompressionQuality = [int]$defaults.manga.LossyCompressionQuality }
    if (@('webp', 'png', 'jpeg', 'avif') -notcontains $cfg.manga.OutputFormat) { $cfg.manga.OutputFormat = [string]$defaults.manga.OutputFormat }
    if ([string]::IsNullOrWhiteSpace($cfg.merge.Language)) { $cfg.merge.Language = [string]$defaults.merge.Language }
    if ([string]::IsNullOrWhiteSpace($cfg.merge.DescriptionHeader)) { $cfg.merge.DescriptionHeader = [string]$defaults.merge.DescriptionHeader }
    if ([string]::IsNullOrWhiteSpace($cfg.merge.MetadataContributor)) { $cfg.merge.MetadataContributor = [string]$defaults.merge.MetadataContributor }
    if (-not ($cfg.manga.ModeScaleSelected -or $cfg.manga.ModeWidthSelected -or $cfg.manga.ModeHeightSelected -or $cfg.manga.ModeFitToDisplaySelected)) {
        $cfg.manga.ModeScaleSelected = $true
    }
    $cfg.kcc.EffectiveBaseArgs = @(Build-KccBaseArgsFromConfig -KccConfig $cfg.kcc)
    $cfg.kcc.EffectiveArgsSource = if ($cfg.kcc.BaseArgs -and $cfg.kcc.BaseArgs.Count -gt 0) { 'kcc.BaseArgs' } else { 'kcc.CliOptions' }

    return [pscustomobject]@{
        ConfigPath = $resolvedPath
        ConfigExists = [bool]$exists
        ConfigValid = [bool]($valid -or -not $exists)
        Config = $cfg
    }
}

function Convert-ToIntBounded {
    param(
        [object]$Value,
        [int]$DefaultValue,
        [int]$MinValue,
        [int]$MaxValue
    )

    if ($null -eq $Value) { return $DefaultValue }
    $parsed = 0
    if (-not [int]::TryParse([string]$Value, [ref]$parsed)) { return $DefaultValue }
    if ($parsed -lt $MinValue) { return $MinValue }
    if ($parsed -gt $MaxValue) { return $MaxValue }
    return $parsed
}

function Resolve-DependenciesConfig {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [string]$DepsConfigPath,
        [string]$KccExePath
    )

    $explicitDepsPath = -not [string]::IsNullOrWhiteSpace($DepsConfigPath)
    $resolvedDepsPath = if ($explicitDepsPath) { $DepsConfigPath } else { Join-Path $ScriptRoot 'manga_epub_automation.deps.json' }
    $configExists = Test-PathSafe -Path $resolvedDepsPath -PathType Leaf

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
    $defaultPaths = $defaults.paths

    $resolvedPaths = [ordered]@{}
    $pathSources = [ordered]@{}
    $envFallbackFields = New-Object 'System.Collections.Generic.List[string]'

    $pathSpecs = @(
        @{ Key = 'python_exe'; Source = 'env_default' },
        @{ Key = 'backend_script'; Source = 'env_default' },
        @{ Key = 'models_dir'; Source = 'env_default' },
        @{ Key = 'kcc_exe'; Source = 'script_default' },
        @{ Key = 'merge_script'; Source = 'script_default' }
    )

    foreach ($spec in $pathSpecs) {
        $key = [string]$spec.Key
        $resolvedValue = ''
        $resolvedSource = 'unresolved'

        if ($key -eq 'kcc_exe' -and -not [string]::IsNullOrWhiteSpace($KccExePath)) {
            $resolvedValue = $KccExePath
            $resolvedSource = 'cli'
        }

        if ([string]::IsNullOrWhiteSpace($resolvedValue) -and $depsConfigValid -and $depsConfig -and ($depsConfig.PSObject.Properties.Name -contains 'paths') -and $depsConfig.paths) {
            if (($depsConfig.paths.PSObject.Properties.Name -contains $key) -and -not [string]::IsNullOrWhiteSpace([string]$depsConfig.paths.$key)) {
                $resolvedValue = [string]$depsConfig.paths.$key
                $resolvedSource = 'deps_json'
            }
        }

        if ([string]::IsNullOrWhiteSpace($resolvedValue) -and ($defaultPaths.PSObject.Properties.Name -contains $key)) {
            $defaultValue = [string]$defaultPaths.$key
            if (-not [string]::IsNullOrWhiteSpace($defaultValue)) {
                $resolvedValue = $defaultValue
                $resolvedSource = [string]$spec.Source
                if ($resolvedSource -eq 'env_default') {
                    $envFallbackFields.Add($key) | Out-Null
                }
            }
        }

        $resolvedPaths[$key] = $resolvedValue
        $pathSources[$key] = $resolvedSource
    }

    $progressDefaults = $defaults.progress
    $progressEnabled = [bool]$progressDefaults.enabled
    $refreshSeconds = [int]$progressDefaults.refresh_seconds
    $etaMinSamples = [int]$progressDefaults.eta_min_samples
    $nonInteractiveInterval = [int]$progressDefaults.noninteractive_log_interval_seconds

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
        PathSources = [pscustomobject]$pathSources
        EnvFallbackFields = @($envFallbackFields)
        ProgressConfig = [pscustomobject]@{
            Enabled = [bool]$progressEnabled
            RefreshSeconds = [int]$refreshSeconds
            EtaMinSamples = [int]$etaMinSamples
            NonInteractiveLogIntervalSeconds = [int]$nonInteractiveInterval
        }
    }
}

function Write-JsonUtf8NoBom {
    param([Parameter(Mandatory = $true)][object]$Object, [Parameter(Mandatory = $true)][string]$Path)
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

function Get-WorkflowRef {
    param([Parameter(Mandatory = $true)][object]$Config)

    if (-not ($Config.PSObject.Properties.Name -contains 'Workflows')) { throw 'Invalid config: missing Workflows' }
    if (-not ($Config.Workflows.PSObject.Properties.Name -contains '$values')) { throw "Invalid config: Workflows missing `'$values`'" }
    if ($Config.Workflows.'$values'.Count -lt 1) { throw 'Invalid config: Workflows.$values is empty' }

    if (-not ($Config.PSObject.Properties.Name -contains 'SelectedWorkflowIndex')) {
        $Config | Add-Member -NotePropertyName SelectedWorkflowIndex -NotePropertyValue 0
    }

    $wfIndex = [int]$Config.SelectedWorkflowIndex
    if ($wfIndex -lt 0 -or $wfIndex -ge $Config.Workflows.'$values'.Count) {
        Write-Warn "SelectedWorkflowIndex=$wfIndex out of range. Fallback to 0."
        $wfIndex = 0
        $Config.SelectedWorkflowIndex = 0
    }
    return $Config.Workflows.'$values'[$wfIndex]
}

function Apply-WorkflowOverrides {
    param(
        [Parameter(Mandatory = $true)][object]$Workflow,
        [Parameter(Mandatory = $true)][object]$Overrides
    )

    if ($null -eq $Overrides) { return }
    if ($Overrides -is [string] -or $Overrides -is [ValueType] -or $Overrides -is [System.Array]) { return }
    if ($null -eq $Overrides.PSObject) { return }
    $overrideProps = @($Overrides.PSObject.Properties)
    if ($overrideProps.Count -lt 1) { return }

    $blocked = @('InputFolderPath', 'OutputFolderPath')
    foreach ($prop in $overrideProps) {
        $name = [string]$prop.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($blocked -contains $name) { continue }
        $value = $prop.Value

        if ($Workflow.PSObject.Properties.Name -contains $name) {
            $Workflow.$name = $value
        }
        else {
            $Workflow | Add-Member -Force -NotePropertyName $name -NotePropertyValue $value
        }
    }
}

function Ensure-BaseSettings {
    param(
        [Parameter(Mandatory = $true)][string]$SettingsPath,
        [Parameter(Mandatory = $true)][string]$RoamingAppStatePath,
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$BackendScript,
        [Parameter(Mandatory = $true)][string]$ModelsDirectory,
        [Parameter(Mandatory = $true)][object]$PipelineConfig
    )

    if (Test-Path -LiteralPath $SettingsPath) { return }
    if (-not (Test-Path -LiteralPath $RoamingAppStatePath)) { throw "Cannot initialize base settings. Missing: $RoamingAppStatePath" }

    $cfg = Read-JsonFile -Path $RoamingAppStatePath
    $cfg | Add-Member -Force -NotePropertyName ScriptMeta -NotePropertyValue ([pscustomobject]@{
        SourceDirSuffixToSkip = [string]$PipelineConfig.pipeline.SourceDirSuffixToSkip
        OutputDirSuffix = [string]$PipelineConfig.pipeline.OutputDirSuffix
        EpubOutputDirSuffix = [string]$PipelineConfig.pipeline.EpubOutputDirSuffix
        OutputFilenamePattern = [string]$PipelineConfig.pipeline.OutputFilenamePattern
    })
    $cfg.ModelsDirectory = ''

    $wf = Get-WorkflowRef -Config $cfg
    $wf.SelectedTabIndex = [int]$PipelineConfig.manga.SelectedTabIndex
    if ($wf.PSObject.Properties.Name -contains 'InputFilePath') { $wf.InputFilePath = '' }
    $wf.InputFolderPath = ''
    $wf.OutputFolderPath = ''
    $wf.OutputFilename = [string]$PipelineConfig.pipeline.OutputFilenamePattern
    $wf.OverwriteExistingFiles = [bool]$PipelineConfig.manga.OverwriteExistingFiles
    $wf.ModeScaleSelected = [bool]$PipelineConfig.manga.ModeScaleSelected
    $wf.ModeWidthSelected = [bool]$PipelineConfig.manga.ModeWidthSelected
    $wf.ModeHeightSelected = [bool]$PipelineConfig.manga.ModeHeightSelected
    if ($wf.PSObject.Properties.Name -contains 'ModeFitToDisplaySelected') { $wf.ModeFitToDisplaySelected = [bool]$PipelineConfig.manga.ModeFitToDisplaySelected }
    $wf.UpscaleScaleFactor = [int]$PipelineConfig.manga.UpscaleScaleFactor
    switch ([string]$PipelineConfig.manga.OutputFormat) {
        'webp' { $wf.WebpSelected = $true;  $wf.PngSelected = $false; $wf.JpegSelected = $false; $wf.AvifSelected = $false }
        'png'  { $wf.WebpSelected = $false; $wf.PngSelected = $true;  $wf.JpegSelected = $false; $wf.AvifSelected = $false }
        'jpeg' { $wf.WebpSelected = $false; $wf.PngSelected = $false; $wf.JpegSelected = $true;  $wf.AvifSelected = $false }
        'avif' { $wf.WebpSelected = $false; $wf.PngSelected = $false; $wf.JpegSelected = $false; $wf.AvifSelected = $true }
    }
    $wf.LossyCompressionQuality = [int]$PipelineConfig.manga.LossyCompressionQuality

    Write-JsonUtf8NoBom -Object $cfg -Path $SettingsPath
    Write-Info "Initialized base settings: $SettingsPath"
}

function Resolve-SourceDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$SkipSuffixes,
        [string]$PreferredName
    )

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { throw "TitleRoot not found or not a directory: $Root" }

    $candidates = @(Get-ChildItem -LiteralPath $Root -Directory | Where-Object {
        $name = $_.Name
        $shouldSkip = $false
        foreach ($suffix in $SkipSuffixes) {
            if ([string]::IsNullOrWhiteSpace($suffix)) { continue }
            if ($name.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $shouldSkip = $true
                break
            }
        }
        -not $shouldSkip
    })

    if ($PreferredName) {
        $selected = @($candidates | Where-Object { $_.Name -eq $PreferredName })
        if ($selected.Count -ne 1) {
            $allNames = if ($candidates.Count -gt 0) { ($candidates | ForEach-Object { $_.Name }) -join ', ' } else { '<none>' }
            throw "SourceDirName '$PreferredName' not found under '$Root'. Candidates: $allNames"
        }
        return $selected[0]
    }

    if ($candidates.Count -eq 1) { return $candidates[0] }
    if ($candidates.Count -eq 0) { throw "No source subdirectory found under '$Root' after suffix filtering: $($SkipSuffixes -join ', ')" }

    $names = $candidates | ForEach-Object { " - $($_.Name)" }
    throw "Multiple source directories found under '$Root'. Use -SourceDirName.`n$($names -join "`n")"
}

function Get-ComicMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$FallbackAuthor,
        [Parameter(Mandatory = $true)][string]$MetadataFileName
    )

    $metaPath = Join-Path $SourceRoot $MetadataFileName
    $comicTitle = Split-Path -Leaf $SourceRoot
    $author = $FallbackAuthor

    if (-not (Test-Path -LiteralPath $metaPath -PathType Leaf)) {
        return [pscustomobject]@{
            ComicTitle = $comicTitle
            Author = $author
            MetadataPath = $null
            MetadataFound = $false
            ParseOk = $false
            TitleFromMetadata = $false
            AuthorFromMetadata = $false
        }
    }

    $parseOk = $false
    $titleFromMeta = $false
    $authorFromMeta = $false

    try {
        $meta = Read-JsonFile -Path $metaPath
        $parseOk = $true
        if ($meta -and ($meta.PSObject.Properties.Name -contains 'comic') -and $meta.comic) {
            $comic = $meta.comic
            if (($comic.PSObject.Properties.Name -contains 'name') -and $comic.name) {
                $comicTitle = [string]$comic.name
                $titleFromMeta = $true
            }
            if (($comic.PSObject.Properties.Name -contains 'author') -and $comic.author -and $comic.author.Count -gt 0) {
                $firstAuthor = $comic.author[0]
                if ($firstAuthor -and ($firstAuthor.PSObject.Properties.Name -contains 'name') -and $firstAuthor.name) {
                    $author = [string]$firstAuthor.name
                    $authorFromMeta = $true
                }
            }
        }
    }
    catch {
        Write-Warn "Failed to parse metadata JSON: $metaPath; using fallback values. Error: $($_.Exception.Message)"
    }

    return [pscustomobject]@{
        ComicTitle = $comicTitle
        Author = $author
        MetadataPath = $metaPath
        MetadataFound = $true
        ParseOk = $parseOk
        TitleFromMetadata = $titleFromMeta
        AuthorFromMetadata = $authorFromMeta
    }
}

function Test-ChapterHasImages {
    param([Parameter(Mandatory = $true)][string]$ChapterPath)
    $allowed = @('.webp', '.png', '.jpg', '.jpeg', '.avif', '.bmp')
    $files = Get-ChildItem -LiteralPath $ChapterPath -File -ErrorAction SilentlyContinue | Where-Object { $allowed -contains $_.Extension.ToLowerInvariant() }
    return $files.Count -gt 0
}

function Get-EpubFileName {
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$ChapterDirName,
        [Parameter(Mandatory = $true)][string]$ComicTitle,
        [Parameter(Mandatory = $true)][string]$TankobonGroupName,
        [Parameter(Mandatory = $true)][string]$VolumeLabel
    )

    if ($GroupName -eq $TankobonGroupName) {
        $match = [regex]::Match($ChapterDirName, '^\u7B2C(?<num>\d+(?:\.\d+)?)\s*[\u5377\u5DFB]$')
        if ($match.Success) { return "{0} - {1}{2}.epub" -f $ComicTitle, $VolumeLabel, $match.Groups['num'].Value }
        return "{0} - {1}.epub" -f $ComicTitle, $ChapterDirName
    }

    return "{0}.epub" -f $ChapterDirName
}

function Get-KccCliSafeText {
    param(
        [string]$Value,
        [string]$Fallback = 'mjnai'
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Fallback }
    $chars = New-Object 'System.Collections.Generic.List[string]'
    foreach ($ch in $Value.ToCharArray()) {
        $code = [int][char]$ch
        if ($code -ge 32 -and $code -le 126) {
            $chars.Add([string]$ch) | Out-Null
        }
        else {
            $chars.Add('_') | Out-Null
        }
    }
    $safe = (($chars.ToArray()) -join '') -replace '\s+', ' '
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return $Fallback }
    return $safe
}

function Repair-EpubDisplayMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$EpubPath,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Author,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    if (-not (Test-Path -LiteralPath $EpubPath -PathType Leaf)) { return }

    $tempPath = Join-Path (Split-Path -Parent $EpubPath) ("mjnai_epubmeta_{0}.epub" -f ([guid]::NewGuid().ToString('N')))
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $escapedAuthor = [System.Security.SecurityElement]::Escape($Author)

    $srcZip = $null
    $dstZip = $null
    try {
        $srcZip = [System.IO.Compression.ZipFile]::OpenRead($EpubPath)
        $dstZip = [System.IO.Compression.ZipFile]::Open($tempPath, [System.IO.Compression.ZipArchiveMode]::Create)

        $opfEntry = $srcZip.Entries | Where-Object { $_.FullName -match '(?i)\.opf$' } | Select-Object -First 1
        $ncxEntry = $srcZip.Entries | Where-Object { $_.FullName -match '(?i)\.ncx$' } | Select-Object -First 1
        $navEntries = @($srcZip.Entries | Where-Object { $_.FullName -match '(?i)(^|/)nav\.xhtml$' } | ForEach-Object { $_.FullName })
        $opfName = if ($opfEntry) { $opfEntry.FullName } else { $null }
        $ncxName = if ($ncxEntry) { $ncxEntry.FullName } else { $null }

        foreach ($entry in $srcZip.Entries) {
            $compressionLevel = if ($entry.FullName -eq 'mimetype') { [System.IO.Compression.CompressionLevel]::NoCompression } else { [System.IO.Compression.CompressionLevel]::Optimal }
            $newEntry = $dstZip.CreateEntry($entry.FullName, $compressionLevel)

            $sourceStream = $entry.Open()
            $targetStream = $newEntry.Open()
            try {
                $isNav = $navEntries -contains $entry.FullName
                if ($entry.FullName -eq $opfName -or $entry.FullName -eq $ncxName -or $isNav) {
                    $reader = [System.IO.StreamReader]::new($sourceStream, [System.Text.Encoding]::UTF8, $true)
                    $text = $reader.ReadToEnd()
                    $reader.Dispose()

                    if ($entry.FullName -eq $opfName) {
                        $text = [regex]::Replace($text, '<dc:title[^>]*>.*?</dc:title>', "<dc:title>$escapedTitle</dc:title>", 'IgnoreCase,Singleline')
                        $text = [regex]::Replace($text, '<dc:creator[^>]*>.*?</dc:creator>', "<dc:creator>$escapedAuthor</dc:creator>", 'IgnoreCase,Singleline')
                    }
                    elseif ($entry.FullName -eq $ncxName) {
                        $text = [regex]::Replace($text, '<docTitle>\s*<text>.*?</text>\s*</docTitle>', "<docTitle><text>$escapedTitle</text></docTitle>", 'IgnoreCase,Singleline')
                        $text = [regex]::Replace($text, '<navLabel>\s*<text>.*?</text>\s*</navLabel>', "<navLabel><text>$escapedTitle</text></navLabel>", 'IgnoreCase,Singleline')
                    }
                    elseif ($isNav) {
                        $text = [regex]::Replace($text, '<title[^>]*>.*?</title>', "<title>$escapedTitle</title>", 'IgnoreCase,Singleline')
                        $text = [regex]::Replace($text, '(<a\b[^>]*>).*?(</a>)', "`$1$escapedTitle`$2", 'IgnoreCase,Singleline')
                    }

                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
                    $targetStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $sourceStream.CopyTo($targetStream)
                }
            }
            finally {
                $targetStream.Dispose()
                $sourceStream.Dispose()
            }
        }

        $dstZip.Dispose()
        $srcZip.Dispose()
        $dstZip = $null
        $srcZip = $null

        Move-Item -LiteralPath $tempPath -Destination $EpubPath -Force
        Write-StageMessage -Message ("EPUB_META_REWRITE: {0}" -f $EpubPath) -LogPath $LogPath -DebugOnly
    }
    catch {
        Write-Warn ("Failed to rewrite EPUB metadata for display: {0} ; {1}" -f $EpubPath, $_.Exception.Message)
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        if ($null -ne $dstZip) { $dstZip.Dispose() }
        if ($null -ne $srcZip) { $srcZip.Dispose() }
    }
}

function Invoke-KccPack {
    param(
        [Parameter(Mandatory = $true)][string]$KccPath,
        [Parameter(Mandatory = $true)][string]$InputChapterPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Author,
        [string]$TargetEpubPath = '',
        [string[]]$KccBaseArgs = @('-p', 'KS', '-f', 'EPUB', '--nokepub', '-n', '--forcecolor'),
        [bool]$EnableUnicodeStaging = $true,
        [string]$UnicodeStagePrefix = 'manga_epub_automation_kcc_stage_',
        [string]$SafeTitleFallback = 'manga_epub_automation_book',
        [string]$SafeAuthorFallback = 'pipeline',
        [bool]$EnableMetadataRewrite = $true,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $finalTargetPath = if ([string]::IsNullOrWhiteSpace($TargetEpubPath)) { Join-Path $OutputDirectory ($Title + '.epub') } else { $TargetEpubPath }
    $hasUnicode = ($InputChapterPath -match '[^\u0000-\u007F]') -or ($OutputDirectory -match '[^\u0000-\u007F]') -or ($Title -match '[^\u0000-\u007F]') -or ($Author -match '[^\u0000-\u007F]') -or ($finalTargetPath -match '[^\u0000-\u007F]')

    if ((-not $hasUnicode) -or (-not $EnableUnicodeStaging)) {
        $kccArgs = @($KccBaseArgs + @('-t', $Title, '-a', $Author, '-o', $OutputDirectory, $InputChapterPath))
        $kccRun = Invoke-ExternalWithLogging -Executable $KccPath -Arguments $kccArgs -LogPath $LogPath
        $kccCode = [int]$kccRun.ExitCode
        if ($EnableMetadataRewrite -and $kccCode -eq 0 -and (Test-Path -LiteralPath $finalTargetPath -PathType Leaf)) {
            Repair-EpubDisplayMetadata -EpubPath $finalTargetPath -Title $Title -Author $Author -LogPath $LogPath
        }
        return $kccCode
    }

    $stagePrefix = if ([string]::IsNullOrWhiteSpace($UnicodeStagePrefix)) { 'manga_epub_automation_kcc_stage_' } else { $UnicodeStagePrefix }
    $stageRoot = Join-Path $env:TEMP ($stagePrefix + [guid]::NewGuid().ToString('N'))
    $stageInput = Join-Path $stageRoot 'input\chapter'
    $stageOutput = Join-Path $stageRoot 'output'
    $safeTitle = Get-KccCliSafeText -Value $Title -Fallback $SafeTitleFallback
    $safeAuthor = Get-KccCliSafeText -Value $Author -Fallback $SafeAuthorFallback

    try {
        New-Item -ItemType Directory -Force -Path $stageInput, $stageOutput | Out-Null
        Get-ChildItem -LiteralPath $InputChapterPath -Force -ErrorAction Stop | Copy-Item -Destination $stageInput -Recurse -Force
        Write-StageMessage -Message ("EPUB_STAGE: unicode-safe staging enabled: {0}" -f $InputChapterPath) -LogPath $LogPath -DebugOnly

        $kccArgs = @($KccBaseArgs + @('-t', $safeTitle, '-a', $safeAuthor, '-o', $stageOutput, $stageInput))
        $kccRun = Invoke-ExternalWithLogging -Executable $KccPath -Arguments $kccArgs -LogPath $LogPath
        $kccCode = [int]$kccRun.ExitCode
        if ($kccCode -ne 0) { return $kccCode }

        $built = @(Get-ChildItem -LiteralPath $stageOutput -File -Filter '*.epub' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        if ($built.Count -lt 1) {
            Write-StageMessage -Message 'EPUB_FAIL: KCC succeeded but no epub found in staging output.' -LogPath $LogPath -Warning
            return 97
        }

        $targetDir = Split-Path -Parent $finalTargetPath
        if (-not [string]::IsNullOrWhiteSpace($targetDir)) {
            New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
        }
        Copy-Item -LiteralPath $built[0].FullName -Destination $finalTargetPath -Force
        if ($EnableMetadataRewrite) {
            Repair-EpubDisplayMetadata -EpubPath $finalTargetPath -Title $Title -Author $Author -LogPath $LogPath
        }
        return 0
    }
    finally {
        try {
            if (Test-Path -LiteralPath $stageRoot) {
                [System.IO.Directory]::Delete($stageRoot, $true)
            }
        }
        catch {
            Write-Info ("KCC staging cleanup skipped: {0}" -f $stageRoot)
        }
    }
}
function Convert-ToNullableDouble {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return [double]$number
    }
    if ([double]::TryParse($text, [ref]$number)) { return [double]$number }
    return $null
}

function TryParse-ChapterOrderFromName {
    param([Parameter(Mandatory = $true)][string]$ChapterName)

    $m = [regex]::Match($ChapterName, '(?<num>\d+(?:\.\d+)?)\s*[\u8BDD\u8A71]')
    if (-not $m.Success) { return $null }
    return Convert-ToNullableDouble -Value $m.Groups['num'].Value
}

function Format-OrderValue {
    param([double]$Value)
    return ([double]$Value).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-StringSha256 {
    param([Parameter(Mandatory = $true)][string]$Value)

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha.ComputeHash($bytes)
        return ([BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-MergePlanManifestHash {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Chapters,
        [Parameter(Mandatory = $true)][string]$TargetName,
        [string]$OrderOverrideSignature = 'AUTO'
    )

    $hashLines = @()
    foreach ($c in @($Chapters)) {
        $orderText = if ($c.HasOrder) { Format-OrderValue -Value ([double]$c.Order) } else { '' }
        $hashLines += ("{0}|{1}|{2}|{3}|{4}" -f $c.ChapterName, $orderText, $c.FileLength, $c.LastWriteTicks, $c.EpubPath)
    }
    $hashLines += ("TARGET={0}" -f $TargetName)
    $hashLines += ("ORDER_OVERRIDE={0}" -f $OrderOverrideSignature)
    return Get-StringSha256 -Value ($hashLines -join "`n")
}

function Get-DefaultChapterOrderMap {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$DefaultGroupName,
        [Parameter(Mandatory = $true)][string]$MetadataFileName,
        [Parameter(Mandatory = $true)][string]$ChapterMetadataFileName
    )

    $map = @{}

    $metaPath = Join-Path $SourceRoot $MetadataFileName
    if (Test-Path -LiteralPath $metaPath -PathType Leaf) {
        try {
            $meta = Read-JsonFile -Path $metaPath
            $defaultEntries = $null
            if ($meta -and ($meta.PSObject.Properties.Name -contains 'comic') -and $meta.comic -and ($meta.comic.PSObject.Properties.Name -contains 'groups') -and $meta.comic.groups) {
                if ($meta.comic.groups.PSObject.Properties.Name -contains 'default') {
                    $defaultEntries = $meta.comic.groups.default
                }
            }

            foreach ($entry in @($defaultEntries)) {
                if ($null -eq $entry) { continue }
                $ord = Convert-ToNullableDouble -Value $entry.order
                if ($null -eq $ord) { continue }

                $chapterName = $null
                if (($entry.PSObject.Properties.Name -contains 'chapterDownloadDir') -and $entry.chapterDownloadDir) {
                    $chapterName = Split-Path -Leaf ([string]$entry.chapterDownloadDir)
                }
                if ([string]::IsNullOrWhiteSpace($chapterName) -and ($entry.PSObject.Properties.Name -contains 'chapterTitle') -and $entry.chapterTitle) {
                    $chapterName = [string]$entry.chapterTitle
                }
                if (-not [string]::IsNullOrWhiteSpace($chapterName)) {
                    $map[$chapterName] = [double]$ord
                }
            }
        }
        catch {
            Write-Warn "Failed parsing root metadata for order map: $metaPath ; $($_.Exception.Message)"
        }
    }

    $defaultSourceDir = Join-Path $SourceRoot $DefaultGroupName
    if (Test-Path -LiteralPath $defaultSourceDir -PathType Container) {
        $chapterDirs = @(Get-ChildItem -LiteralPath $defaultSourceDir -Directory -ErrorAction SilentlyContinue)
        foreach ($chapterDir in $chapterDirs) {
            $chapterMetaPath = Join-Path $chapterDir.FullName $ChapterMetadataFileName
            if (-not (Test-Path -LiteralPath $chapterMetaPath -PathType Leaf)) { continue }

            try {
                $chapterMeta = Read-JsonFile -Path $chapterMetaPath
                $ord = Convert-ToNullableDouble -Value $chapterMeta.order
                if ($null -ne $ord) { $map[$chapterDir.Name] = [double]$ord }
            }
            catch {
                Write-Warn "Failed parsing chapter metadata: $chapterMetaPath ; $($_.Exception.Message)"
            }
        }
    }

    return $map
}

function Get-MergedEpubPlan {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$UpscaledDefaultGroupPath,
        [Parameter(Mandatory = $true)][string]$MergedOutputDefaultGroupPath,
        [Parameter(Mandatory = $true)][string]$ComicTitle,
        [Parameter(Mandatory = $true)][string]$DefaultGroupName,
        [Parameter(Mandatory = $true)][string]$MetadataFileName,
        [Parameter(Mandatory = $true)][string]$ChapterMetadataFileName,
        [Parameter(Mandatory = $true)][string]$TalkLabel
    )

    if (-not (Test-Path -LiteralPath $UpscaledDefaultGroupPath -PathType Container)) {
        return [pscustomobject]@{ Chapters=@(); TargetPath=$null; TargetFileName=$null; ManifestHash=$null; MergedFiles=@() }
    }

    $orderMap = Get-DefaultChapterOrderMap -SourceRoot $SourceRoot -DefaultGroupName $DefaultGroupName -MetadataFileName $MetadataFileName -ChapterMetadataFileName $ChapterMetadataFileName
    $mergedNamePattern = '^{0}\s*-\s*{1}.+$' -f [regex]::Escape($ComicTitle), [regex]::Escape($TalkLabel)

    $epubFiles = @(Get-ChildItem -LiteralPath $UpscaledDefaultGroupPath -File -Filter '*.epub' -ErrorAction SilentlyContinue | Sort-Object Name)
    $mergedOutputFiles = if (Test-Path -LiteralPath $MergedOutputDefaultGroupPath -PathType Container) {
        @(Get-ChildItem -LiteralPath $MergedOutputDefaultGroupPath -File -Filter '*.epub' -ErrorAction SilentlyContinue | Where-Object {
                [regex]::IsMatch([System.IO.Path]::GetFileNameWithoutExtension($_.Name), $mergedNamePattern)
            })
    }
    else {
        @()
    }
    $mergedFiles = @($mergedOutputFiles)

    $chapterCandidates = @()
    foreach ($epub in $epubFiles) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($epub.Name)
        $chapterDirPath = Join-Path $UpscaledDefaultGroupPath $base
        $isChapterDir = Test-Path -LiteralPath $chapterDirPath -PathType Container
        $hasOrderFromMap = $orderMap.ContainsKey($base)

        if (-not $isChapterDir -and -not $hasOrderFromMap) { continue }

        $ord = if ($hasOrderFromMap) { [double]$orderMap[$base] } else { TryParse-ChapterOrderFromName -ChapterName $base }
        $chapterCandidates += [pscustomobject]@{
            ChapterName = $base
            EpubPath = $epub.FullName
            HasOrder = ($null -ne $ord)
            Order = $ord
            FileLength = [int64]$epub.Length
            LastWriteTicks = [int64]$epub.LastWriteTimeUtc.Ticks
        }
    }

    if ($chapterCandidates.Count -eq 0) {
        return [pscustomobject]@{ Chapters=@(); TargetPath=$null; TargetFileName=$null; ManifestHash=$null; MergedFiles=$mergedFiles }
    }

    $sorted = @($chapterCandidates | Sort-Object `
            @{ Expression = { if ($_.HasOrder) { 0 } else { 1 } } }, `
            @{ Expression = { if ($_.HasOrder) { [double]$_.Order } else { [double]::PositiveInfinity } } }, `
            @{ Expression = { $_.ChapterName } })

    $orderedOnly = @($sorted | Where-Object { $_.HasOrder })
    if ($orderedOnly.Count -gt 0) {
        $startText = Format-OrderValue -Value ([double]$orderedOnly[0].Order)
        $endText = Format-OrderValue -Value ([double]$orderedOnly[$orderedOnly.Count - 1].Order)
        $rangeLabel = if ($startText -eq $endText) { "{0}{1}" -f $TalkLabel, $startText } else { "{0}{1}-{2}" -f $TalkLabel, $startText, $endText }
    }
    else {
        $rangeLabel = "{0}合集" -f $TalkLabel
    }

    $targetName = "{0} - {1}.epub" -f $ComicTitle, $rangeLabel
    $targetPath = Join-Path $MergedOutputDefaultGroupPath $targetName

    $manifestHash = Get-MergePlanManifestHash -Chapters $sorted -TargetName $targetName -OrderOverrideSignature 'AUTO'

    return [pscustomobject]@{
        Chapters = $sorted
        TargetPath = $targetPath
        TargetFileName = $targetName
        ManifestHash = $manifestHash
        MergedFiles = $mergedFiles
    }
}

function Get-MergeOrderOverrideState {
    param([Parameter(Mandatory = $true)][string]$OrderFilePath)

    $exists = Test-Path -LiteralPath $OrderFilePath -PathType Leaf
    $state = [pscustomobject]@{
        Path = $OrderFilePath
        Exists = [bool]$exists
        IsValid = $false
        Errors = @()
        Warnings = @()
        RawNames = @()
        OrderNames = @()
        Signature = 'AUTO'
        Applied = $false
    }

    if (-not $exists) { return $state }

    $doc = $null
    try {
        $doc = Read-JsonFile -Path $OrderFilePath
    }
    catch {
        $state.Errors += ("Invalid merge order JSON: {0}" -f $_.Exception.Message)
        return $state
    }

    if (-not ($doc.PSObject.Properties.Name -contains 'chapters')) {
        $state.Errors += 'Merge order file missing required field: chapters'
        return $state
    }

    $chaptersValue = $doc.chapters
    if ($null -eq $chaptersValue) {
        $state.Errors += 'Merge order field chapters cannot be null.'
        return $state
    }
    if ($chaptersValue -is [string]) {
        $state.Errors += 'Merge order field chapters must be an array of strings.'
        return $state
    }
    if (-not ($chaptersValue -is [System.Array] -or $chaptersValue -is [System.Collections.IList])) {
        $state.Errors += 'Merge order field chapters must be an array of strings.'
        return $state
    }

    $rawList = New-Object 'System.Collections.Generic.List[string]'
    $orderList = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($entry in @($chaptersValue)) {
        if ($null -eq $entry) {
            $state.Warnings += 'Merge order entry is null and has been ignored.'
            continue
        }
        $name = ([string]$entry).Trim()
        [void]$rawList.Add($name)
        if ([string]::IsNullOrWhiteSpace($name)) {
            $state.Warnings += 'Merge order entry is empty after trim and has been ignored.'
            continue
        }
        if (-not $seen.Add($name)) {
            $state.Warnings += ("Duplicate merge order entry ignored (first one kept): {0}" -f $name)
            continue
        }
        [void]$orderList.Add($name)
    }

    $sigLines = @()
    $versionText = if ($doc.PSObject.Properties.Name -contains 'version') { [string]$doc.version } else { '' }
    $sigLines += ("version={0}" -f $versionText)
    foreach ($name in $rawList) { $sigLines += ("chapter={0}" -f $name) }

    $state.RawNames = @($rawList.ToArray())
    $state.OrderNames = @($orderList.ToArray())
    $state.Signature = Get-StringSha256 -Value ($sigLines -join "`n")
    $state.IsValid = $true
    return $state
}

function Apply-MergeOrderOverrideToPlan {
    param(
        [Parameter(Mandatory = $true)][object]$MergePlan,
        [Parameter(Mandatory = $true)][string]$OrderFilePath
    )

    $state = Get-MergeOrderOverrideState -OrderFilePath $OrderFilePath
    $chapters = @($MergePlan.Chapters)

    if ($state.Exists -and $state.IsValid -and $chapters.Count -gt 0) {
        $chapterMap = @{}
        foreach ($chapter in $chapters) {
            $key = ([string]$chapter.ChapterName).Trim()
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            if (-not $chapterMap.ContainsKey($key)) {
                $chapterMap[$key] = New-Object 'System.Collections.Generic.List[object]'
            }
                [void]$chapterMap[$key].Add($chapter)
        }

        $ordered = New-Object 'System.Collections.Generic.List[object]'
        $usedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $unmatched = New-Object 'System.Collections.Generic.List[string]'

        foreach ($name in @($state.OrderNames)) {
            if (-not $chapterMap.ContainsKey($name)) {
                [void]$unmatched.Add($name)
                continue
            }
            foreach ($chapter in $chapterMap[$name]) {
                $uniqueKey = [string]$chapter.EpubPath
                if ($usedPaths.Add($uniqueKey)) {
                    [void]$ordered.Add($chapter)
                }
            }
        }

        foreach ($chapter in $chapters) {
            $uniqueKey = [string]$chapter.EpubPath
            if ($usedPaths.Add($uniqueKey)) {
                [void]$ordered.Add($chapter)
            }
        }

        if ($unmatched.Count -gt 0) {
            $state.Warnings += ("Merge order entries not found in current candidates and ignored: {0}" -f (($unmatched.ToArray()) -join ', '))
        }
        $state.Applied = $true
        $chapters = @($ordered.ToArray())
    }

    $targetName = if (-not [string]::IsNullOrWhiteSpace([string]$MergePlan.TargetFileName)) {
        [string]$MergePlan.TargetFileName
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$MergePlan.TargetPath)) {
        [System.IO.Path]::GetFileName([string]$MergePlan.TargetPath)
    }
    else {
        ''
    }

    $signature = if ($state.Exists -and $state.IsValid) { [string]$state.Signature } else { 'AUTO' }
    $manifestHash = if ([string]::IsNullOrWhiteSpace($targetName)) {
        [string]$MergePlan.ManifestHash
    }
    else {
        Get-MergePlanManifestHash -Chapters $chapters -TargetName $targetName -OrderOverrideSignature $signature
    }

    $plan = [pscustomobject]@{
        Chapters = $chapters
        TargetPath = $MergePlan.TargetPath
        TargetFileName = $MergePlan.TargetFileName
        ManifestHash = $manifestHash
        MergedFiles = $MergePlan.MergedFiles
    }

    return [pscustomobject]@{
        Plan = $plan
        State = $state
    }
}

function Export-MergeOrderTemplate {
    param(
        [Parameter(Mandatory = $true)][string]$OrderFilePath,
        [Parameter(Mandatory = $true)][object]$MergePlan
    )

    $orderDir = Split-Path -Parent $OrderFilePath
    if (-not [string]::IsNullOrWhiteSpace($orderDir)) {
        New-Item -ItemType Directory -Force -Path $orderDir | Out-Null
    }

    $obj = [pscustomobject]@{
        version = 1
        chapters = @($MergePlan.Chapters | ForEach-Object { [string]$_.ChapterName })
    }
    Write-JsonUtf8NoBom -Object $obj -Path $OrderFilePath
    return [pscustomobject]@{
        Path = $OrderFilePath
        ChapterCount = $obj.chapters.Count
    }
}

function Get-MergedRebuildDecision {
    param(
        [Parameter(Mandatory = $true)][string]$ManifestPath,
        [Parameter(Mandatory = $true)][string]$ManifestHash,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$MergedFiles
    )

    $otherMerged = @($MergedFiles | Where-Object { $_.FullName -ne $TargetPath })
    if ($otherMerged.Count -gt 0) { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'old-merged-exists'; OtherMerged = $otherMerged } }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'manifest-missing'; OtherMerged = $otherMerged } }

    try { $manifest = Read-JsonFile -Path $ManifestPath }
    catch { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'manifest-invalid'; OtherMerged = $otherMerged } }

    if (-not ($manifest.PSObject.Properties.Name -contains 'ManifestHash')) { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'manifest-no-hash'; OtherMerged = $otherMerged } }
    if ([string]$manifest.ManifestHash -ne $ManifestHash) { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'manifest-hash-changed'; OtherMerged = $otherMerged } }
    if (-not (Test-Path -LiteralPath $TargetPath -PathType Leaf)) { return [pscustomobject]@{ NeedRebuild = $true; Reason = 'target-missing'; OtherMerged = $otherMerged } }

    return [pscustomobject]@{ NeedRebuild = $false; Reason = 'up-to-date'; OtherMerged = $otherMerged }
}

function Invoke-MergedEpubPack {
    param(
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$MergeScriptPath,
        [Parameter(Mandatory = $true)][string]$PlanPath,
        [Parameter(Mandatory = $true)][string]$LogPath
)

    $mergeRun = Invoke-ExternalWithLogging -Executable $PythonExe -Arguments @($MergeScriptPath, '--plan', $PlanPath) -LogPath $LogPath
    return [int]$mergeRun.ExitCode
}

function Add-PreflightIssue {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
        [Parameter(Mandatory = $true)][ValidateSet('ERROR', 'WARN', 'INFO')][string]$Severity,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $Issues.Add([pscustomobject]@{
            Severity = $Severity
            Code = $Code
            Message = $Message
        }) | Out-Null
}

function Write-RunResultFile {
    param(
        [Parameter(Mandatory = $true)][string]$ResultPath,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string]$PlanPath,
        [Parameter(Mandatory = $true)][object]$Summary
    )

    $obj = [pscustomobject]@{
        Status = $Status
        Message = $Message
        CreatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        PlanPath = $PlanPath
        Summary = $Summary
    }
    Write-JsonUtf8NoBom -Object $obj -Path $ResultPath
}

function Write-RunArtifacts {
    param(
        [Parameter(Mandatory = $true)][string]$ResultPath,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string]$PlanPath,
        [Parameter(Mandatory = $true)][object]$Summary
    )

    Write-RunResultFile -ResultPath $ResultPath -Status $Status -Message $Message -PlanPath $PlanPath -Summary $Summary
    if (-not [string]::IsNullOrWhiteSpace($script:LatestRunResultPath)) {
        Update-LatestArtifact -SourcePath $ResultPath -LatestPath $script:LatestRunResultPath
    }
    Write-GuiEvent -Type 'run_result' -Data ([pscustomobject]@{
            status = $Status
            message = $Message
            plan_path = $PlanPath
            latest_plan_path = $script:LatestRunPlanPath
            result_path = $ResultPath
            latest_result_path = $script:LatestRunResultPath
        })
}

function Build-ExecutionPlan {
    param([Parameter(Mandatory = $true)][object]$PlanSummary)
    return $PlanSummary
}

function Run-PreflightGate {
    param([Parameter(Mandatory = $true)][scriptblock]$GateBody)
    . $GateBody
}

function Invoke-EpubStage {
    param(
        [Parameter(Mandatory = $true)][bool]$ShouldRun,
        [Parameter(Mandatory = $true)][scriptblock]$StageBody
    )
    if ($ShouldRun) { . $StageBody }
}

function Invoke-MergeStage {
    param(
        [Parameter(Mandatory = $true)][bool]$ShouldRun,
        [Parameter(Mandatory = $true)][scriptblock]$StageBody
    )
    if ($ShouldRun) { . $StageBody }
}

function Get-OutputImageExtension {
    param([Parameter(Mandatory = $true)][string]$Format)
    switch ($Format) {
        'webp' { return '.webp' }
        'png' { return '.png' }
        'jpeg' { return '.jpg' }
        'avif' { return '.avif' }
    }
    return '.webp'
}

function Test-DirectoryWritable {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    try {
        if (-not (Test-Path -LiteralPath $DirectoryPath -PathType Container)) {
            return [pscustomobject]@{ IsWritable = $false; Error = "Directory not found: $DirectoryPath" }
        }
        $probe = Join-Path $DirectoryPath (".manga_epub_automation_probe_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
        [System.IO.File]::WriteAllText($probe, 'probe', [System.Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ IsWritable = $true; Error = $null }
    }
    catch {
        return [pscustomobject]@{ IsWritable = $false; Error = $_.Exception.Message }
    }
}

function Format-SecondsAsHms {
    param([double]$Seconds)
    if ($Seconds -lt 0) { $Seconds = 0 }
    $ts = [System.TimeSpan]::FromSeconds([math]::Floor($Seconds))
    return ('{0:00}:{1:00}:{2:00}' -f [int]$ts.Hours, [int]$ts.Minutes, [int]$ts.Seconds)
}

function Get-UpscaleProgressSnapshot {
    param(
        [int]$TotalUnits,
        [int]$DoneUnits,
        [int]$ProcessedUnits,
        [int]$SkippedUnits,
        [double]$ElapsedSeconds,
        [int]$EtaMinSamples
    )

    $effectiveTotal = [math]::Max([double]$TotalUnits, [double]$DoneUnits)
    if ($effectiveTotal -le 0) {
        return [pscustomobject]@{
            Percent = 100
            EffectiveTotal = 0
            AvgRate = 0.0
            EtaSeconds = $null
            StatusText = 'done=0/0 processed=0 skipped=0 ETA=<n/a> rate=0.00 img/s'
        }
    }

    $percent = [math]::Floor((100.0 * [double]$DoneUnits) / $effectiveTotal)
    if ($percent -gt 100) { $percent = 100 }
    if ($percent -lt 0) { $percent = 0 }

    $safeElapsed = [math]::Max(0.001, $ElapsedSeconds)
    $rate = [double]$DoneUnits / $safeElapsed
    $eta = $null
    if ($DoneUnits -ge $EtaMinSamples -and $rate -gt 0.00001) {
        $remaining = [math]::Max(0.0, $effectiveTotal - [double]$DoneUnits)
        $eta = $remaining / $rate
    }
    $etaText = if ($null -eq $eta) { '<estimating>' } else { Format-SecondsAsHms -Seconds $eta }

    return [pscustomobject]@{
        Percent = [int]$percent
        EffectiveTotal = [int]$effectiveTotal
        AvgRate = [double]$rate
        EtaSeconds = $eta
        StatusText = ("done={0}/{1} processed={2} skipped={3} ETA={4} rate={5:N2} img/s" -f $DoneUnits, [int]$effectiveTotal, $ProcessedUnits, $SkippedUnits, $etaText, $rate)
    }
}

function Get-EpubPackProgressSnapshot {
    param(
        [int]$DoneUnits,
        [int]$TotalUnits,
        [int]$PackedUnits,
        [int]$SkippedUnits,
        [int]$FailedUnits,
        [double]$ElapsedSeconds
    )

    $safeTotal = [math]::Max(0, $TotalUnits)
    if ($safeTotal -le 0) {
        return [pscustomobject]@{
            Percent = 100
            StatusText = 'done=0/0 packed=0 skipped=0 failed=0 ETA=<n/a>'
            EtaText = '<n/a>'
        }
    }

    $safeDone = [math]::Max(0, [math]::Min($DoneUnits, $safeTotal))
    $percent = [math]::Floor((100.0 * $safeDone) / $safeTotal)
    $etaText = '<estimating>'
    $safeElapsed = [math]::Max(0.001, $ElapsedSeconds)
    $rate = [double]$safeDone / $safeElapsed
    if ($safeDone -gt 0 -and $rate -gt 0.00001) {
        $remaining = [math]::Max(0.0, [double]$safeTotal - [double]$safeDone)
        $etaText = Format-SecondsAsHms -Seconds ($remaining / $rate)
    }

    return [pscustomobject]@{
        Percent = [int]$percent
        StatusText = ("done={0}/{1} packed={2} skipped={3} failed={4} ETA={5}" -f $safeDone, $safeTotal, $PackedUnits, $SkippedUnits, $FailedUnits, $etaText)
        EtaText = $etaText
    }
}

function Invoke-UpscaleStage {
    param(
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$BackendScript,
        [Parameter(Mandatory = $true)][string]$RuntimeSettingsPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][int]$TotalUnits,
        [Parameter(Mandatory = $true)][object]$ProgressConfig,
        [Parameter(Mandatory = $true)][string]$LogLevel,
        [switch]$NoUpscaleProgress
    )

    $allLines = New-Object 'System.Collections.Generic.List[string]'
    $processedCount = 0
    $skipCount = 0
    $doneCount = 0
    $errorLineCount = 0

    $progressEnabled = ([bool]$ProgressConfig.Enabled) -and (-not $NoUpscaleProgress)
    $refreshSeconds = [int]$ProgressConfig.RefreshSeconds
    if ($refreshSeconds -lt 1) { $refreshSeconds = 1 }
    $etaMinSamples = [int]$ProgressConfig.EtaMinSamples
    if ($etaMinSamples -lt 1) { $etaMinSamples = 1 }
    $nonInteractiveInterval = [int]$ProgressConfig.NonInteractiveLogIntervalSeconds
    if ($nonInteractiveInterval -lt 1) { $nonInteractiveInterval = 10 }

    $interactive = [Environment]::UserInteractive
    $lastUiUpdate = [datetime]::MinValue
    $lastTextUpdate = [datetime]::MinValue

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $logWriter = [System.IO.StreamWriter]::new($LogPath, $true, [System.Text.UTF8Encoding]::new($false))
    try {
        & $PythonExe $BackendScript --settings $RuntimeSettingsPath 2>&1 | ForEach-Object {
            $line = [string]$_
            $allLines.Add($line)
            $logWriter.WriteLine($line)
            $logWriter.Flush()

            if ($line -like 'PROGRESS=*') {
                $processedCount += 1
                $doneCount += 1
            }
            elseif ($line -like '*file exists, skip:*') {
                $skipCount += 1
                $doneCount += 1
            }

            $isErrorLine = $line -match '(?i)(traceback|exception|\berror\b)'
            if ($isErrorLine) { $errorLineCount += 1 }

            if ($LogLevel -eq 'debug') {
                Write-Host $line
            }
            elseif ($isErrorLine) {
                Write-Host $line -ForegroundColor Yellow
            }

            if ($progressEnabled) {
                $now = Get-Date
                $snapshot = Get-UpscaleProgressSnapshot -TotalUnits $TotalUnits -DoneUnits $doneCount -ProcessedUnits $processedCount -SkippedUnits $skipCount -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds -EtaMinSamples $etaMinSamples
                $emitGuiProgress = $false
                if ($interactive) {
                    if (($now - $lastUiUpdate).TotalSeconds -ge $refreshSeconds) {
                        Write-Progress -Id 1 -Activity 'Upscaling Images' -Status $snapshot.StatusText -PercentComplete $snapshot.Percent
                        $lastUiUpdate = $now
                        $emitGuiProgress = $true
                    }
                }
                elseif (($now - $lastTextUpdate).TotalSeconds -ge $nonInteractiveInterval) {
                    Write-Host ("UPSCALE_PROGRESS: {0}% {1}" -f $snapshot.Percent, $snapshot.StatusText)
                    $lastTextUpdate = $now
                    $emitGuiProgress = $true
                }
                if ($script:PipelineGuiMode -and $emitGuiProgress) {
                    Write-GuiEvent -Type 'upscale_progress' -Data ([pscustomobject]@{
                            percent = [int]$snapshot.Percent
                            done_units = [int]$doneCount
                            total_units = [int]$snapshot.EffectiveTotal
                            processed_units = [int]$processedCount
                            skipped_units = [int]$skipCount
                            eta_seconds = if ($null -eq $snapshot.EtaSeconds) { $null } else { [double]$snapshot.EtaSeconds }
                            avg_rate_img_per_sec = [double]$snapshot.AvgRate
                            elapsed_seconds = [double]$stopwatch.Elapsed.TotalSeconds
                            status = $snapshot.StatusText
                        })
                }
            }
        }
        $exitCode = $LASTEXITCODE
    }
    finally {
        $stopwatch.Stop()
        $logWriter.Dispose()
        if ($progressEnabled -and $interactive) {
            Write-Progress -Id 1 -Activity 'Upscaling Images' -Completed
        }
    }

    $finalSnapshot = Get-UpscaleProgressSnapshot -TotalUnits $TotalUnits -DoneUnits $doneCount -ProcessedUnits $processedCount -SkippedUnits $skipCount -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds -EtaMinSamples $etaMinSamples
    if ($progressEnabled) {
        Write-Host ("UPSCALE_PROGRESS: {0}% {1}" -f $finalSnapshot.Percent, $finalSnapshot.StatusText)
    }
    if ($script:PipelineGuiMode) {
        Write-GuiEvent -Type 'upscale_progress' -Data ([pscustomobject]@{
                percent = [int]$finalSnapshot.Percent
                done_units = [int]$doneCount
                total_units = [int]$finalSnapshot.EffectiveTotal
                processed_units = [int]$processedCount
                skipped_units = [int]$skipCount
                eta_seconds = if ($null -eq $finalSnapshot.EtaSeconds) { $null } else { [double]$finalSnapshot.EtaSeconds }
                avg_rate_img_per_sec = [double]$finalSnapshot.AvgRate
                elapsed_seconds = [double]$stopwatch.Elapsed.TotalSeconds
                status = $finalSnapshot.StatusText
            })
    }

    return [pscustomobject]@{
        ExitCode = [int]$exitCode
        Lines = @($allLines)
        ProcessedCount = [int]$processedCount
        SkipCount = [int]$skipCount
        DoneCount = [int]$doneCount
        ErrorLineCount = [int]$errorLineCount
        ElapsedSeconds = [double]$stopwatch.Elapsed.TotalSeconds
        AvgRateImgPerSec = [double]$finalSnapshot.AvgRate
        ProgressSummary = [pscustomobject]@{
            total_units = [int][math]::Max($TotalUnits, $doneCount)
            done_units = [int]$doneCount
            processed_units = [int]$processedCount
            skipped_units = [int]$skipCount
            avg_rate_img_per_sec = [double]$finalSnapshot.AvgRate
            elapsed_seconds = [double]$stopwatch.Elapsed.TotalSeconds
        }
    }
}

function Estimate-UpscaleWork {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$OutputRoot,
        [Parameter(Mandatory = $true)][string]$OutputExt,
        [Parameter(Mandatory = $true)][string]$FileNamePattern
    )

    $allowed = @('.webp', '.png', '.jpg', '.jpeg', '.avif', '.bmp')
    $inputFiles = @(Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
            $allowed -contains $_.Extension.ToLowerInvariant()
        })
    $processable = 0
    $predSkip = 0

    foreach ($file in $inputFiles) {
        $rel = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $relDir = Split-Path -Parent $rel
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $outName = if ($FileNamePattern -match '%filename%') {
            ($FileNamePattern -replace '%filename%', $name)
        }
        else {
            "$name-upscaled"
        }
        $predPath = if ([string]::IsNullOrWhiteSpace($relDir)) {
            Join-Path $OutputRoot ($outName + $OutputExt)
        }
        else {
            Join-Path (Join-Path $OutputRoot $relDir) ($outName + $OutputExt)
        }
        $processable += 1
        if (Test-Path -LiteralPath $predPath -PathType Leaf) {
            $predSkip += 1
        }
    }

    return [pscustomobject]@{
        InputImages = $inputFiles.Count
        PredictedProcess = ($processable - $predSkip)
        PredictedSkip = $predSkip
    }
}

function Estimate-EpubPackWork {
    param(
        [Parameter(Mandatory = $true)][string]$UpscaledRoot,
        [Parameter(Mandatory = $true)][string]$PublishedOutputRoot,
        [Parameter(Mandatory = $true)][string]$ComicTitle,
        [Parameter(Mandatory = $true)][string]$TankobonGroupName,
        [Parameter(Mandatory = $true)][string]$DefaultGroupName,
        [Parameter(Mandatory = $true)][string]$VolumeLabel
    )

    $planned = 0
    $skipped = 0
    $groups = @($TankobonGroupName, $DefaultGroupName)
    foreach ($group in $groups) {
        $sourceGroupPath = Join-Path $UpscaledRoot $group
        if (-not (Test-Path -LiteralPath $sourceGroupPath -PathType Container)) { continue }
        $targetGroupPath = if ($group -eq $TankobonGroupName) { $PublishedOutputRoot } else { $sourceGroupPath }
        $chapterDirs = @(Get-ChildItem -LiteralPath $sourceGroupPath -Directory -ErrorAction SilentlyContinue)
        foreach ($chapterDir in $chapterDirs) {
            if (-not (Test-ChapterHasImages -ChapterPath $chapterDir.FullName)) { continue }
            $epubName = Get-EpubFileName -GroupName $group -ChapterDirName $chapterDir.Name -ComicTitle $ComicTitle -TankobonGroupName $TankobonGroupName -VolumeLabel $VolumeLabel
            $targetPath = Join-Path $targetGroupPath $epubName
            $planned += 1
            if (Test-Path -LiteralPath $targetPath -PathType Leaf) { $skipped += 1 }
        }
    }
    return [pscustomobject]@{ Planned = $planned; PredictedProcess = ($planned - $skipped); PredictedSkip = $skipped }
}

function Show-MergePreview {
    param(
        [Parameter(Mandatory = $true)][object]$MergePlan,
        [Parameter(Mandatory = $true)][bool]$Compact
    )

    if ($MergePlan.Chapters.Count -eq 0) {
        Write-Info '[PLAN] MERGE order: no chapter candidates'
        Write-GuiEvent -Type 'merge_preview' -Data ([pscustomobject]@{
                target_path = $MergePlan.TargetPath
                chapter_count = 0
                compact = [bool]$Compact
                chapters = @()
            })
        return
    }

    $rows = @()
    $idx = 1
    foreach ($ch in $MergePlan.Chapters) {
        $rows += [pscustomobject]@{
            Index = $idx
            Chapter = $ch.ChapterName
            Order = if ($ch.HasOrder) { Format-OrderValue -Value ([double]$ch.Order) } else { '<none>' }
            EpubPath = $ch.EpubPath
        }
        $idx += 1
    }

    Write-GuiEvent -Type 'merge_preview' -Data ([pscustomobject]@{
            target_path = $MergePlan.TargetPath
            chapter_count = $MergePlan.Chapters.Count
            compact = [bool]$Compact
            chapters = @($rows)
        })

    Write-Host ''
    Write-Host '========== Merge Order Preview =========='
    Write-Host ("Target: {0}" -f $MergePlan.TargetPath)
    Write-Host ("Chapters: {0}" -f $MergePlan.Chapters.Count)
    if ($Compact -and $rows.Count -gt 30) {
        $head = $rows | Select-Object -First 10
        $tail = $rows | Select-Object -Last 10
        $head | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
        Write-Host ("... {0} rows omitted ..." -f ($rows.Count - 20))
        $tail | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
    }
    else {
        $rows | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
    }

    $missing = @($MergePlan.Chapters | Where-Object { -not $_.HasOrder } | ForEach-Object { $_.ChapterName })
    if ($missing.Count -gt 0) { Write-Warn ("Missing order chapters: {0}" -f ($missing -join ', ')) }
    $dups = @($MergePlan.Chapters | Where-Object { $_.HasOrder } | Group-Object { Format-OrderValue -Value ([double]$_.Order) } | Where-Object { $_.Count -gt 1 })
    if ($dups.Count -gt 0) {
        $dupText = $dups | ForEach-Object { "{0}({1})" -f $_.Name, $_.Count }
        Write-Warn ("Duplicate order values: {0}" -f ($dupText -join ', '))
    }
    $nonStandard = @($MergePlan.Chapters | Where-Object { $_.ChapterName -notmatch '\d+(?:\.\d+)?\s*[\u8BDD\u8A71]' } | ForEach-Object { $_.ChapterName })
    if ($nonStandard.Count -gt 0) { Write-Warn ("Non-standard chapter names: {0}" -f ($nonStandard -join ', ')) }
    Write-Host '========================================='
}
if ($Help) {
    Get-Help -Full $PSCommandPath
    exit 0
}

$scriptRoot = Split-Path -Parent $PSCommandPath
$processingRoot = $scriptRoot
$logsDir = Join-Path $processingRoot 'logs'
$settingsPath = Join-Path $processingRoot 'manga_epub_automation.settings.json'
$runtimeSettingsPath = Join-Path $processingRoot 'manga_epub_automation.runtime.json'
$mergeOrderFileName = '.manga_epub_automation_merge_order.json'
$mergeManifestFileName = '.manga_epub_automation_merge_manifest.json'

if ([string]::IsNullOrWhiteSpace($DepsConfigPath)) {
    $DepsConfigPath = Join-Path $scriptRoot 'manga_epub_automation.deps.json'
}
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $scriptRoot 'manga_epub_automation.config.json'
}

if ($InitConfig) {
    $configTemplate = Get-DefaultPipelineConfigObject
    $configDir = Split-Path -Parent $ConfigPath
    if (-not [string]::IsNullOrWhiteSpace($configDir)) {
        New-Item -ItemType Directory -Force -Path $configDir | Out-Null
    }
    if (Test-Path -LiteralPath $ConfigPath -PathType Leaf) {
        Write-Warn "Pipeline config already exists, keep unchanged: $ConfigPath"
    }
    else {
        Write-JsonUtf8NoBom -Object $configTemplate -Path $ConfigPath
        Write-Info "Initialized pipeline config: $ConfigPath"
    }
    exit 0
}
if ($InitDepsConfig) {
    $depsTemplate = Get-DefaultDepsConfigObject -ScriptRoot $scriptRoot
    $depsDir = Split-Path -Parent $DepsConfigPath
    if (-not [string]::IsNullOrWhiteSpace($depsDir)) {
        New-Item -ItemType Directory -Force -Path $depsDir | Out-Null
    }
    if (Test-Path -LiteralPath $DepsConfigPath -PathType Leaf) {
        Write-Warn "Deps config already exists, keep unchanged: $DepsConfigPath"
    }
    else {
        Write-JsonUtf8NoBom -Object $depsTemplate -Path $DepsConfigPath
        Write-Info "Initialized deps config: $DepsConfigPath"
    }
    exit 0
}

if ([string]::IsNullOrWhiteSpace($TitleRoot)) {
    throw 'Parameter -TitleRoot is required unless -Help, -InitConfig, or -InitDepsConfig is used.'
}

if ([string]::IsNullOrWhiteSpace($MergedEpubAuthorFallback)) {
    $MergedEpubAuthorFallback = $EpubAuthorFallback
}

$runUpscaleStage = -not $SkipUpscale
$runEpubStage = -not $SkipEpubPackaging
$runMergeStage = -not $SkipMergedEpub

$depsResolution = Resolve-DependenciesConfig -ScriptRoot $scriptRoot -DepsConfigPath $DepsConfigPath -KccExePath $KccExePath
$pipelineConfigResolution = Resolve-PipelineConfig -ScriptRoot $scriptRoot -ConfigPath $ConfigPath
$pipelineConfig = $pipelineConfigResolution.Config
$pythonExe = [string]$depsResolution.ResolvedPaths.python_exe
$backendScript = [string]$depsResolution.ResolvedPaths.backend_script
$modelsDirectory = [string]$depsResolution.ResolvedPaths.models_dir
$KccExePath = [string]$depsResolution.ResolvedPaths.kcc_exe
$mergeScriptPath = [string]$depsResolution.ResolvedPaths.merge_script
$progressConfig = $depsResolution.ProgressConfig

$roamingAppStatePath = if ($env:APPDATA) { Join-Path $env:APPDATA 'MangaJaNaiConverterGui\appstate2.json' } else { '' }

New-Item -ItemType Directory -Force -Path $logsDir | Out-Null

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $logsDir ("manga_epub_automation_run_{0}.log" -f $timestamp)
$planPath = Join-Path $logsDir ("run_plan_{0}.json" -f $timestamp)
$resultPath = Join-Path $logsDir ("run_result_{0}.json" -f $timestamp)
$script:LatestRunPlanPath = Join-Path $logsDir 'latest_run_plan.json'
$script:LatestRunResultPath = Join-Path $logsDir 'latest_run_result.json'

$sourceSuffixToSkip = [string]$pipelineConfig.pipeline.SourceDirSuffixToSkip
$outputDirSuffix = [string]$pipelineConfig.pipeline.OutputDirSuffix
$epubOutputDirSuffix = [string]$pipelineConfig.pipeline.EpubOutputDirSuffix
$outputFilenamePattern = [string]$pipelineConfig.pipeline.OutputFilenamePattern
if (-not $pipelineConfigResolution.ConfigExists -and (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
    try {
        $existingCfg = Read-JsonFile -Path $settingsPath
        if ($existingCfg.PSObject.Properties.Name -contains 'ScriptMeta') {
            if (($existingCfg.ScriptMeta.PSObject.Properties.Name -contains 'SourceDirSuffixToSkip') -and $existingCfg.ScriptMeta.SourceDirSuffixToSkip) { $sourceSuffixToSkip = [string]$existingCfg.ScriptMeta.SourceDirSuffixToSkip }
            if (($existingCfg.ScriptMeta.PSObject.Properties.Name -contains 'OutputDirSuffix') -and $existingCfg.ScriptMeta.OutputDirSuffix) { $outputDirSuffix = [string]$existingCfg.ScriptMeta.OutputDirSuffix }
            if (($existingCfg.ScriptMeta.PSObject.Properties.Name -contains 'EpubOutputDirSuffix') -and $existingCfg.ScriptMeta.EpubOutputDirSuffix) { $epubOutputDirSuffix = [string]$existingCfg.ScriptMeta.EpubOutputDirSuffix }
            if (($existingCfg.ScriptMeta.PSObject.Properties.Name -contains 'OutputFilenamePattern') -and $existingCfg.ScriptMeta.OutputFilenamePattern) { $outputFilenamePattern = [string]$existingCfg.ScriptMeta.OutputFilenamePattern }
        }
    }
    catch {
        Write-Warn "Failed to read existing settings for script meta defaults: $($_.Exception.Message)"
    }
}

$effectiveUpscaleFactor = if ($PSBoundParameters.ContainsKey('UpscaleFactor')) { [int]$UpscaleFactor } else { [int]$pipelineConfig.manga.UpscaleScaleFactor }
$effectiveOutputFormat = if ($PSBoundParameters.ContainsKey('OutputFormat')) { ([string]$OutputFormat).ToLowerInvariant() } else { [string]$pipelineConfig.manga.OutputFormat }
$effectiveLossyQuality = if ($PSBoundParameters.ContainsKey('LossyQuality')) { [int]$LossyQuality } else { [int]$pipelineConfig.manga.LossyCompressionQuality }

$sourceDir = $null
$sourceResolveError = $null
try {
    $sourceDir = Resolve-SourceDirectory -Root $TitleRoot -SkipSuffixes @($sourceSuffixToSkip, $epubOutputDirSuffix) -PreferredName $SourceDirName
}
catch {
    $sourceResolveError = $_.Exception.Message
}

$outputDir = if ($null -ne $sourceDir) { Join-Path $TitleRoot ($sourceDir.Name + $outputDirSuffix) } else { Join-Path $TitleRoot ('<unresolved-source>' + $outputDirSuffix) }
$epubOutputDir = if ($null -ne $sourceDir) { Join-Path $TitleRoot ($sourceDir.Name + $epubOutputDirSuffix) } else { Join-Path $TitleRoot ('<unresolved-source>' + $epubOutputDirSuffix) }
$epubOutputDefaultGroupPath = $epubOutputDir
if ([string]::IsNullOrWhiteSpace($MergeOrderFilePath)) {
    $MergeOrderFilePath = Join-Path $epubOutputDir $mergeOrderFileName
}
$comicMeta = if ($null -ne $sourceDir) {
    Get-ComicMetadata -SourceRoot $sourceDir.FullName -FallbackAuthor $EpubAuthorFallback -MetadataFileName $MetadataJsonName
}
else {
    [pscustomobject]@{
        ComicTitle = '<unresolved>'
        Author = $EpubAuthorFallback
        MetadataPath = $null
        MetadataFound = $false
        ParseOk = $false
        TitleFromMetadata = $false
        AuthorFromMetadata = $false
    }
}

$hasPython = Test-PathSafe -Path $pythonExe -PathType Leaf
$hasBackend = Test-PathSafe -Path $backendScript -PathType Leaf
$hasModels = Test-PathSafe -Path $modelsDirectory -PathType Container
$hasKcc = Test-PathSafe -Path $KccExePath -PathType Leaf
$hasMergeScript = Test-PathSafe -Path $mergeScriptPath -PathType Leaf

$mergePlan = [pscustomobject]@{ Chapters=@(); TargetPath=$null; TargetFileName=$null; ManifestHash=$null; MergedFiles=@() }
$mergeOrderState = [pscustomobject]@{ Path = $MergeOrderFilePath; Exists = $false; IsValid = $false; Errors = @(); Warnings = @(); RawNames = @(); OrderNames = @(); Signature = 'AUTO'; Applied = $false }
if ($runMergeStage -and $null -ne $sourceDir) {
    $defaultGroupPathForPlan = Join-Path $outputDir $GroupDefault
    $mergePlanRaw = Get-MergedEpubPlan -SourceRoot $sourceDir.FullName -UpscaledDefaultGroupPath $defaultGroupPathForPlan -MergedOutputDefaultGroupPath $epubOutputDefaultGroupPath -ComicTitle $comicMeta.ComicTitle -DefaultGroupName $GroupDefault -MetadataFileName $MetadataJsonName -ChapterMetadataFileName $ChapterMetadataJsonName -TalkLabel $TalkChar
    $mergeOrderApplied = Apply-MergeOrderOverrideToPlan -MergePlan $mergePlanRaw -OrderFilePath $MergeOrderFilePath
    $mergePlan = $mergeOrderApplied.Plan
    $mergeOrderState = $mergeOrderApplied.State
}
$mergeDecision = if ($mergePlan.Chapters.Count -gt 0) {
    $manifestPathForPlan = Join-Path $epubOutputDefaultGroupPath $mergeManifestFileName
    Get-MergedRebuildDecision -ManifestPath $manifestPathForPlan -ManifestHash $mergePlan.ManifestHash -TargetPath $mergePlan.TargetPath -MergedFiles $mergePlan.MergedFiles
}
else {
    [pscustomobject]@{ NeedRebuild = $false; Reason = 'no-chapters'; OtherMerged = @() }
}

if ($DumpMergeOrderTemplate) {
    if ($null -eq $sourceDir) {
        throw "Cannot dump merge order template because source directory cannot be resolved: $sourceResolveError"
    }
    $defaultGroupPathForTemplate = Join-Path $outputDir $GroupDefault
    $templatePlan = Get-MergedEpubPlan -SourceRoot $sourceDir.FullName -UpscaledDefaultGroupPath $defaultGroupPathForTemplate -MergedOutputDefaultGroupPath $epubOutputDefaultGroupPath -ComicTitle $comicMeta.ComicTitle -DefaultGroupName $GroupDefault -MetadataFileName $MetadataJsonName -ChapterMetadataFileName $ChapterMetadataJsonName -TalkLabel $TalkChar
    $dumpInfo = Export-MergeOrderTemplate -OrderFilePath $MergeOrderFilePath -MergePlan $templatePlan
    Write-Info ("Merge order template exported: {0} (chapters={1})" -f $dumpInfo.Path, $dumpInfo.ChapterCount)
    exit 0
}

$upscaleEstimate = if ($runUpscaleStage -and $null -ne $sourceDir) {
    Estimate-UpscaleWork -SourceRoot $sourceDir.FullName -OutputRoot $outputDir -OutputExt (Get-OutputImageExtension -Format $effectiveOutputFormat) -FileNamePattern $outputFilenamePattern
}
else {
    [pscustomobject]@{ InputImages = 0; PredictedProcess = 0; PredictedSkip = 0 }
}

$epubEstimate = if ($runEpubStage -and $null -ne $sourceDir) {
    Estimate-EpubPackWork -UpscaledRoot $outputDir -PublishedOutputRoot $epubOutputDir -ComicTitle $comicMeta.ComicTitle -TankobonGroupName $GroupTankobon -DefaultGroupName $GroupDefault -VolumeLabel $VolumeChar
}
else {
    [pscustomobject]@{ Planned = 0; PredictedProcess = 0; PredictedSkip = 0 }
}

$issues = New-Object 'System.Collections.Generic.List[object]'
if (-not $depsResolution.ConfigExists) {
    if ($depsResolution.ConfigPathExplicit) {
        Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_CONFIG_MISSING' -Message "Deps config not found: $($depsResolution.ConfigPath)"
    }
    else {
        Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'DEPS_CONFIG_MISSING' -Message ("Deps config not found at default path, fallback chain is in use: {0}" -f $depsResolution.ConfigPath)
    }
}
if ($depsResolution.ConfigExists -and -not $depsResolution.ConfigValid) {
    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_CONFIG_INVALID_JSON' -Message ("Failed to parse deps config JSON: {0} ; {1}" -f $depsResolution.ConfigPath, $depsResolution.ConfigError)
}
if ($depsResolution.EnvFallbackFields.Count -gt 0) {
    Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'DEPS_ENV_FALLBACK_USED' -Message ("Resolved from env defaults: {0}" -f ($depsResolution.EnvFallbackFields -join ', '))
}

if ($runUpscaleStage -or $runMergeStage) {
    if (-not $hasPython) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_PATH_PYTHON' -Message "Python not found: $pythonExe" }
}
if ($runUpscaleStage) {
    if (-not $hasBackend) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_PATH_BACKEND' -Message "Backend script not found: $backendScript" }
    if (-not $hasModels) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_PATH_MODELS' -Message "Models directory not found: $modelsDirectory" }
}
if ($runEpubStage -and -not $hasKcc) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_PATH_KCC' -Message "KCC executable not found: $KccExePath" }
if ($runMergeStage -and -not $hasMergeScript) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'DEPS_PATH_MERGE' -Message "Merge script not found: $mergeScriptPath" }
if ($sourceResolveError) { Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'SOURCE_RESOLVE' -Message $sourceResolveError }

if (-not ($runUpscaleStage -or $runEpubStage -or $runMergeStage)) {
    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'STAGE_NONE_SELECTED' -Message 'No stage selected. Remove at least one -Skip* switch to run work.'
}
if ($runUpscaleStage -and (-not $runEpubStage) -and $runMergeStage) {
    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'STAGE_COMBO_INVALID' -Message 'Invalid stage combination: upscale + merge without EPUB packaging in between.'
}

$probeDir = if (Test-Path -LiteralPath $outputDir -PathType Container) { $outputDir } else { $TitleRoot }
$writable = Test-DirectoryWritable -DirectoryPath $probeDir
if (-not $writable.IsWritable) {
    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'OUTPUT_WRITE' -Message "Output probe directory not writable: $probeDir ; $($writable.Error)"
}
$publishProbeDir = if (Test-Path -LiteralPath $epubOutputDir -PathType Container) { $epubOutputDir } else { $TitleRoot }
$publishWritable = Test-DirectoryWritable -DirectoryPath $publishProbeDir
if (-not $publishWritable.IsWritable) {
    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'PUBLISH_OUTPUT_WRITE' -Message "Publish output probe directory not writable: $publishProbeDir ; $($publishWritable.Error)"
}

if ($null -ne $sourceDir) {
    if (-not $comicMeta.MetadataFound) { Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'META_MISSING' -Message "Missing $MetadataJsonName under source root. Title/author will fallback." }
    elseif (-not $comicMeta.ParseOk) { Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'META_PARSE' -Message "Failed to parse $MetadataJsonName cleanly. Title/author may fallback." }
    elseif (-not $comicMeta.TitleFromMetadata -or -not $comicMeta.AuthorFromMetadata) { Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'META_PARTIAL' -Message 'Metadata available but title/author not fully resolved from metadata.' }

    $sourceGroups = @($GroupTankobon, $GroupDefault)
    foreach ($group in $sourceGroups) {
        $groupPath = Join-Path $sourceDir.FullName $group
        if (-not (Test-Path -LiteralPath $groupPath -PathType Container)) { continue }
        $chapterDirs = @(Get-ChildItem -LiteralPath $groupPath -Directory -ErrorAction SilentlyContinue)
        foreach ($chapterDir in $chapterDirs) {
            if (-not (Test-ChapterHasImages -ChapterPath $chapterDir.FullName)) {
                Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'CHAPTER_NO_IMAGE' -Message "No image files in source chapter folder: $($chapterDir.FullName)"
            }
        }
    }
}

if ($runMergeStage) {
    if ($mergeOrderState.Exists -and -not $mergeOrderState.IsValid) {
        Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'MERGE_ORDER_FILE_INVALID' -Message ("Invalid merge order file: {0} ; {1}" -f $MergeOrderFilePath, (($mergeOrderState.Errors) -join ' | '))
    }
    if ($mergeOrderState.Exists -and $mergeOrderState.IsValid) {
        foreach ($msg in @($mergeOrderState.Warnings)) {
            Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'MERGE_ORDER_FILE_WARN' -Message $msg
        }
    }

    if ($mergePlan.Chapters.Count -eq 0) {
        if ($runEpubStage) {
            Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'MERGE_NO_CHAPTERS_YET' -Message 'Merge enabled but no chapter EPUB candidates found at preflight. This is expected before EPUB stage on first/incremental runs.'
        }
        else {
            Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'MERGE_NO_CHAPTERS' -Message 'Merge enabled but no chapter EPUB candidates found in default upscaled group.'
        }
    }
    else {
        $missingOrder = @($mergePlan.Chapters | Where-Object { -not $_.HasOrder })
        if ($missingOrder.Count -gt 0) {
            Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'MERGE_ORDER_MISSING' -Message ("{0} merge chapter(s) missing order and will sort to tail: {1}" -f $missingOrder.Count, (($missingOrder | ForEach-Object { $_.ChapterName }) -join ', '))
        }
        $dupOrders = @($mergePlan.Chapters | Where-Object { $_.HasOrder } | Group-Object { Format-OrderValue -Value ([double]$_.Order) } | Where-Object { $_.Count -gt 1 })
        if ($dupOrders.Count -gt 0) {
            Add-PreflightIssue -Issues $issues -Severity 'WARN' -Code 'MERGE_ORDER_DUPLICATE' -Message ("Duplicate merge order values: {0}" -f (($dupOrders | ForEach-Object { "{0}({1})" -f $_.Name, $_.Count }) -join ', '))
        }
        foreach ($chapter in $mergePlan.Chapters) {
            if (-not (Test-Path -LiteralPath $chapter.EpubPath -PathType Leaf)) {
                Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'MERGE_EPUB_MISSING' -Message "Merge input epub missing: $($chapter.EpubPath)"
            }
            else {
                try {
                    $fs = [System.IO.File]::OpenRead($chapter.EpubPath)
                    $fs.Dispose()
                }
                catch {
                    Add-PreflightIssue -Issues $issues -Severity 'ERROR' -Code 'MERGE_EPUB_UNREADABLE' -Message "Merge input epub unreadable: $($chapter.EpubPath) ; $($_.Exception.Message)"
                }
            }
        }
        if ($PlanOnly -or $LogLevel -eq 'debug') {
            Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'MERGE_REBUILD' -Message ("Merge decision: rebuild={0} reason={1}" -f $mergeDecision.NeedRebuild, $mergeDecision.Reason)
            if ($mergeDecision.OtherMerged.Count -gt 0) {
                Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'MERGE_DELETE_OLD' -Message ("Old merged epubs scheduled for replacement: {0}" -f (($mergeDecision.OtherMerged | ForEach-Object { $_.Name }) -join ', '))
            }
        }
    }
}

if ($LogLevel -eq 'debug') {
    if ($runUpscaleStage) {
        Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'UPSCALE_SKIP_PRED' -Message ("Predicted upscale skip: {0} / {1}" -f $upscaleEstimate.PredictedSkip, $upscaleEstimate.InputImages)
    }
    if ($runEpubStage) {
        Add-PreflightIssue -Issues $issues -Severity 'INFO' -Code 'EPUB_SKIP_PRED' -Message ("Predicted epub skip: {0} / {1}" -f $epubEstimate.PredictedSkip, $epubEstimate.Planned)
    }
}

$issueArray = [object[]]$issues.ToArray()
$errorCount = @($issueArray | Where-Object { $_.Severity -eq 'ERROR' }).Count
$warnCount = @($issueArray | Where-Object { $_.Severity -eq 'WARN' }).Count
$infoCount = @($issueArray | Where-Object { $_.Severity -eq 'INFO' }).Count

$planSummary = [pscustomobject]@{
    TitleRoot = $TitleRoot
    SourceDir = if ($null -ne $sourceDir) { $sourceDir.FullName } else { $null }
    OutputDir = $outputDir
    EpubOutputDir = $epubOutputDir
    Modes = [pscustomobject]@{
        DryRun = [bool]$DryRun
        PlanOnly = [bool]$PlanOnly
        AutoConfirm = [bool]$AutoConfirm
        LogLevel = $LogLevel
        NoUpscaleProgress = [bool]$NoUpscaleProgress
        SkipUpscale = [bool]$SkipUpscale
        SkipEpubPackaging = [bool]$SkipEpubPackaging
        SkipMergedEpub = [bool]$SkipMergedEpub
        MergePreviewCompact = [bool]$MergePreviewCompact
        MergeOrderFilePath = $MergeOrderFilePath
        DumpMergeOrderTemplate = [bool]$DumpMergeOrderTemplate
        FailOnPreflightWarnings = [bool]$FailOnPreflightWarnings
    }
    StageSelection = [pscustomobject]@{
        RunUpscale = [bool]$runUpscaleStage
        RunEpubPackaging = [bool]$runEpubStage
        RunMergedEpub = [bool]$runMergeStage
    }
    ResolvedDependencies = [pscustomobject]@{
        DepsConfigPath = $depsResolution.ConfigPath
        DepsConfigExists = [bool]$depsResolution.ConfigExists
        DepsConfigValid = [bool]$depsResolution.ConfigValid
        PathSources = $depsResolution.PathSources
        Paths = [pscustomobject]@{
            PythonExe = $pythonExe
            BackendScript = $backendScript
            ModelsDirectory = $modelsDirectory
            KccExe = $KccExePath
            MergeScript = $mergeScriptPath
        }
    }
    ResolvedPipelineConfig = [pscustomobject]@{
        ConfigPath = $pipelineConfigResolution.ConfigPath
        ConfigExists = [bool]$pipelineConfigResolution.ConfigExists
        ConfigValid = [bool]$pipelineConfigResolution.ConfigValid
        Effective = [pscustomobject]@{
            SourceDirSuffixToSkip = $sourceSuffixToSkip
            OutputDirSuffix = $outputDirSuffix
            EpubOutputDirSuffix = $epubOutputDirSuffix
            OutputFilenamePattern = $outputFilenamePattern
            UpscaleScaleFactor = [int]$effectiveUpscaleFactor
            OutputFormat = $effectiveOutputFormat
            LossyCompressionQuality = [int]$effectiveLossyQuality
            WorkflowOverrides = $pipelineConfig.manga.WorkflowOverrides
            KccBaseArgs = @($pipelineConfig.kcc.EffectiveBaseArgs)
            KccArgsSource = $pipelineConfig.kcc.EffectiveArgsSource
            KccCliOptions = $pipelineConfig.kcc.CliOptions
            KccUnicodeStaging = $pipelineConfig.kcc.UnicodeStaging
            KccMetadataRewrite = $pipelineConfig.kcc.MetadataRewrite
            MergeLanguage = $pipelineConfig.merge.Language
            MergeDescriptionHeader = $pipelineConfig.merge.DescriptionHeader
            MergeIncludeOrderInDescription = [bool]$pipelineConfig.merge.IncludeOrderInDescription
            MergeMetadataContributor = $pipelineConfig.merge.MetadataContributor
        }
    }
    ProgressConfig = $progressConfig
    UpscaleProgressEstimate = [pscustomobject]@{
        total_units = [int]$upscaleEstimate.InputImages
        initial_skip_pred = [int]$upscaleEstimate.PredictedSkip
    }
    Predicted = [pscustomobject]@{
        Upscale = $upscaleEstimate
        EpubPackaging = $epubEstimate
        Merge = [pscustomobject]@{
            ChapterCount = $mergePlan.Chapters.Count
            TargetPath = $mergePlan.TargetPath
            Rebuild = $mergeDecision.NeedRebuild
            RebuildReason = $mergeDecision.Reason
            ReplaceOldFiles = @($mergeDecision.OtherMerged | ForEach-Object { $_.FullName })
            OrderFilePath = $MergeOrderFilePath
            OrderFileExists = [bool]$mergeOrderState.Exists
            OrderFileApplied = [bool]$mergeOrderState.Applied
            OrderFileWarnings = @($mergeOrderState.Warnings)
        }
    }
    Preflight = [pscustomobject]@{
        ErrorCount = $errorCount
        WarnCount = $warnCount
        InfoCount = $infoCount
        Issues = $issueArray
    }
}
$planSummary = Build-ExecutionPlan -PlanSummary $planSummary
Write-JsonUtf8NoBom -Object $planSummary -Path $planPath
if (-not [string]::IsNullOrWhiteSpace($script:LatestRunPlanPath)) {
    Update-LatestArtifact -SourcePath $planPath -LatestPath $script:LatestRunPlanPath
}
Write-GuiEvent -Type 'plan_ready' -Data ([pscustomobject]@{
        plan_path = $planPath
        latest_plan_path = $script:LatestRunPlanPath
        title_root = $TitleRoot
        source_dir = $planSummary.SourceDir
        output_dir = $planSummary.OutputDir
        epub_output_dir = $planSummary.EpubOutputDir
        stages = $planSummary.StageSelection
    })

Write-Host ''
Write-Host '========== Execution Plan =========='
Write-Host ("TitleRoot:    {0}" -f $TitleRoot)
Write-Host ("SourceDir:    {0}" -f $planSummary.SourceDir)
Write-Host ("OutputDir:    {0}" -f $outputDir)
Write-Host ("EpubOutDir:   {0}" -f $epubOutputDir)
Write-Host ("Mode:         DryRun={0} PlanOnly={1} AutoConfirm={2} LogLevel={3}" -f $DryRun, $PlanOnly, $AutoConfirm, $LogLevel)
Write-Host ("Stages:       Upscale={0} Epub={1} Merge={2}" -f $runUpscaleStage, $runEpubStage, $runMergeStage)
Write-Host ("DepsConfig:   {0} (exists={1} valid={2})" -f $depsResolution.ConfigPath, $depsResolution.ConfigExists, $depsResolution.ConfigValid)
Write-Host ("PipelineCfg:  {0} (exists={1} valid={2})" -f $pipelineConfigResolution.ConfigPath, $pipelineConfigResolution.ConfigExists, $pipelineConfigResolution.ConfigValid)
if ($runUpscaleStage) {
    Write-Host 'Upscale:      forecast will be printed at stage start'
}
else {
    Write-Host 'Upscale:      skipped'
}
if ($runEpubStage) {
    Write-Host 'EPUB:         forecast will be printed at stage start'
}
else {
    Write-Host 'EPUB:         skipped'
}
if ($runMergeStage) {
    Write-Host ("MergeOrder:   {0}" -f $MergeOrderFilePath)
    if ($PlanOnly) {
        Write-Host ("Merge:        candidates={0} target={1}" -f $mergePlan.Chapters.Count, $mergePlan.TargetPath)
        Write-Host ("MergeRebuild: {0} ({1})" -f $mergeDecision.NeedRebuild, $mergeDecision.Reason)
    }
    else {
        Write-Host 'Merge:        forecast and order preview will be printed at stage start'
    }
}
else {
    Write-Host 'Merge:        skipped'
}
Write-Host ("Plan file:    {0}" -f $planPath)
Write-Host '===================================='

$errorIssues = @($issueArray | Where-Object { $_.Severity -eq 'ERROR' })
$warnIssues = @($issueArray | Where-Object { $_.Severity -eq 'WARN' })
$infoIssues = @($issueArray | Where-Object { $_.Severity -eq 'INFO' })
Write-GuiEvent -Type 'preflight_summary' -Data ([pscustomobject]@{
        errors = $errorIssues.Count
        warnings = $warnIssues.Count
        infos = $infoIssues.Count
        issues = $issueArray
    })
if ($infoIssues.Count -gt 0) { foreach ($it in $infoIssues) { Write-Info ("[{0}] {1}" -f $it.Code, $it.Message) } }
if ($warnIssues.Count -gt 0) { foreach ($it in $warnIssues) { Write-Warn ("[{0}] {1}" -f $it.Code, $it.Message) } }
if ($errorIssues.Count -gt 0) { foreach ($it in $errorIssues) { Write-Host ("[ERROR] [{0}] {1}" -f $it.Code, $it.Message) -ForegroundColor Red } }

if ($runMergeStage -and $PlanOnly) {
    Show-MergePreview -MergePlan $mergePlan -Compact:$MergePreviewCompact
}

Run-PreflightGate -GateBody {
    if ($errorIssues.Count -gt 0) {
        Write-RunArtifacts -ResultPath $resultPath -Status 'preflight_error' -Message ("Preflight errors: {0}" -f $errorIssues.Count) -PlanPath $planPath -Summary $planSummary
        throw ("Preflight failed with {0} error(s). Fix issues before execution." -f $errorIssues.Count)
    }
    if ($FailOnPreflightWarnings -and $warnIssues.Count -gt 0) {
        Write-RunArtifacts -ResultPath $resultPath -Status 'preflight_warning_blocked' -Message ("Warnings blocked by strict mode: {0}" -f $warnIssues.Count) -PlanPath $planPath -Summary $planSummary
        throw ("Preflight strict mode blocked execution due to {0} warning(s)." -f $warnIssues.Count)
    }

    if ($PlanOnly) {
        Write-RunArtifacts -ResultPath $resultPath -Status 'plan_only' -Message 'Plan generated only.' -PlanPath $planPath -Summary $planSummary
        exit 0
    }

    if (-not $DryRun) {
        if (-not $AutoConfirm) {
            if (-not [Environment]::UserInteractive) {
                Write-GuiEvent -Type 'confirmation_required' -Data ([pscustomobject]@{
                        scope = 'initial_execution'
                        interactive_required = $true
                        message = 'Interactive confirmation required. Re-run with -AutoConfirm in non-interactive mode.'
                    })
                Write-RunArtifacts -ResultPath $resultPath -Status 'needs_confirmation' -Message 'Interactive confirmation required. Re-run with -AutoConfirm in non-interactive mode.' -PlanPath $planPath -Summary $planSummary
                throw 'Interactive confirmation required. Re-run with -AutoConfirm for scheduled/non-interactive execution.'
            }
            foreach ($warn in $warnIssues) {
                Write-GuiEvent -Type 'confirmation_required' -Data ([pscustomobject]@{
                        scope = 'preflight_warning'
                        code = $warn.Code
                        message = $warn.Message
                        default = 'N'
                    })
                $answer = Read-Host ("WARN [{0}] {1}`nContinue? [y/N]" -f $warn.Code, $warn.Message)
                Write-GuiEvent -Type 'confirmation_response' -Data ([pscustomobject]@{
                        scope = 'preflight_warning'
                        code = $warn.Code
                        response = $answer
                        accepted = ($answer.ToUpperInvariant() -eq 'Y')
                    })
                if ($answer.ToUpperInvariant() -ne 'Y') {
                    Write-RunArtifacts -ResultPath $resultPath -Status 'aborted_warning' -Message ("User rejected warning: {0}" -f $warn.Code) -PlanPath $planPath -Summary $planSummary
                    throw ("Execution cancelled by user on warning [{0}]." -f $warn.Code)
                }
            }
            Write-GuiEvent -Type 'confirmation_required' -Data ([pscustomobject]@{
                    scope = 'initial_execution'
                    message = 'Execute pipeline now?'
                    default = 'N'
                })
            $finalConfirm = Read-Host 'Execute pipeline now? [y/N]'
            Write-GuiEvent -Type 'confirmation_response' -Data ([pscustomobject]@{
                    scope = 'initial_execution'
                    response = $finalConfirm
                    accepted = ($finalConfirm.ToUpperInvariant() -eq 'Y')
                })
            if ($finalConfirm.ToUpperInvariant() -ne 'Y') {
                Write-RunArtifacts -ResultPath $resultPath -Status 'aborted_confirm' -Message 'User cancelled at initial execution confirmation.' -PlanPath $planPath -Summary $planSummary
                throw 'Execution cancelled: initial confirmation not approved.'
            }
        }
        else {
            Write-Info 'AutoConfirm enabled. Skipping interactive confirmation prompts.'
        }
    }
}

if ($runUpscaleStage) {
    Ensure-BaseSettings -SettingsPath $settingsPath -RoamingAppStatePath $roamingAppStatePath -PythonExe $pythonExe -BackendScript $backendScript -ModelsDirectory $modelsDirectory -PipelineConfig $pipelineConfig

    $baseConfig = Read-JsonFile -Path $settingsPath
    $workflow = Get-WorkflowRef -Config $baseConfig
    $baseConfig.ModelsDirectory = $modelsDirectory
    $baseConfig | Add-Member -Force -NotePropertyName ScriptMeta -NotePropertyValue ([pscustomobject]@{
            SourceDirSuffixToSkip = $sourceSuffixToSkip
            OutputDirSuffix = $outputDirSuffix
            EpubOutputDirSuffix = $epubOutputDirSuffix
            OutputFilenamePattern = $outputFilenamePattern
    })

    $workflow.SelectedTabIndex = [int]$pipelineConfig.manga.SelectedTabIndex
    if ($workflow.PSObject.Properties.Name -contains 'InputFilePath') { $workflow.InputFilePath = '' }
    $workflow.InputFolderPath = $sourceDir.FullName
    $workflow.OutputFolderPath = $outputDir
    $workflow.OutputFilename = $outputFilenamePattern
    $workflow.OverwriteExistingFiles = [bool]$pipelineConfig.manga.OverwriteExistingFiles
    $workflow.ModeScaleSelected = [bool]$pipelineConfig.manga.ModeScaleSelected
    $workflow.ModeWidthSelected = [bool]$pipelineConfig.manga.ModeWidthSelected
    $workflow.ModeHeightSelected = [bool]$pipelineConfig.manga.ModeHeightSelected
    if ($workflow.PSObject.Properties.Name -contains 'ModeFitToDisplaySelected') { $workflow.ModeFitToDisplaySelected = [bool]$pipelineConfig.manga.ModeFitToDisplaySelected }
    $workflow.UpscaleScaleFactor = [int]$effectiveUpscaleFactor
    $workflow.LossyCompressionQuality = [int]$effectiveLossyQuality
    switch ($effectiveOutputFormat) {
        'webp' { $workflow.WebpSelected = $true;  $workflow.PngSelected = $false; $workflow.JpegSelected = $false; $workflow.AvifSelected = $false }
        'png'  { $workflow.WebpSelected = $false; $workflow.PngSelected = $true;  $workflow.JpegSelected = $false; $workflow.AvifSelected = $false }
        'jpeg' { $workflow.WebpSelected = $false; $workflow.PngSelected = $false; $workflow.JpegSelected = $true;  $workflow.AvifSelected = $false }
        'avif' { $workflow.WebpSelected = $false; $workflow.PngSelected = $false; $workflow.JpegSelected = $false; $workflow.AvifSelected = $true }
    }
    Apply-WorkflowOverrides -Workflow $workflow -Overrides $pipelineConfig.manga.WorkflowOverrides
    Write-JsonUtf8NoBom -Object $baseConfig -Path $runtimeSettingsPath
}

Write-Info "TitleRoot: $TitleRoot"
Write-Info "SourceDir: $($sourceDir.FullName)"
Write-Info "OutputDir: $outputDir"
Write-Info "EpubOutDir: $epubOutputDir"
Write-Info "Runtime settings: $runtimeSettingsPath"
Write-Info "Deps config: $($depsResolution.ConfigPath)"
Write-Info ("Pipeline config: {0} (exists={1} valid={2})" -f $pipelineConfigResolution.ConfigPath, $pipelineConfigResolution.ConfigExists, $pipelineConfigResolution.ConfigValid)
Write-Info ("Effective manga defaults: scale={0} format={1} quality={2} output_pattern={3}" -f $effectiveUpscaleFactor, $effectiveOutputFormat, $effectiveLossyQuality, $outputFilenamePattern)
Write-Info ("Effective KCC base args: {0}" -f (($pipelineConfig.kcc.EffectiveBaseArgs) -join ' '))
Write-Info ("KCC args source: {0}; unicode_staging={1}; metadata_rewrite={2}" -f $pipelineConfig.kcc.EffectiveArgsSource, [bool]$pipelineConfig.kcc.UnicodeStaging.Enabled, [bool]$pipelineConfig.kcc.MetadataRewrite.Enabled)
Write-Info ("Merge metadata config: language={0}; include_order={1}; contributor={2}" -f $pipelineConfig.merge.Language, [bool]$pipelineConfig.merge.IncludeOrderInDescription, $pipelineConfig.merge.MetadataContributor)
Write-Info ("Progress config: enabled={0} refresh={1}s eta_min_samples={2} noninteractive_interval={3}s cli_disable={4} log_level={5}" -f $progressConfig.Enabled, $progressConfig.RefreshSeconds, $progressConfig.EtaMinSamples, $progressConfig.NonInteractiveLogIntervalSeconds, [bool]$NoUpscaleProgress, $LogLevel)
if ($runUpscaleStage) { Write-Info "Backend command: `"$pythonExe`" `"$backendScript`" --settings `"$runtimeSettingsPath`"" } else { Write-Info 'Upscale stage skipped by configuration.' }
if ($runEpubStage) { Write-Info "KCC executable: $KccExePath" } else { Write-Info 'EPUB packaging stage skipped by configuration.' }
if ($runMergeStage) { Write-Info "Merge script: $mergeScriptPath" } else { Write-Info 'Merged EPUB stage skipped by configuration.' }
Write-GuiEvent -Type 'execution_start' -Data ([pscustomobject]@{
        dry_run = [bool]$DryRun
        stage_selection = $planSummary.StageSelection
        log_path = $logPath
        plan_path = $planPath
        latest_plan_path = $script:LatestRunPlanPath
        result_path = $resultPath
        latest_result_path = $script:LatestRunResultPath
    })

if ($DryRun) {
    Write-Info 'DryRun enabled. Preflight and plan were shown; no processing executed.'
    Write-RunArtifacts -ResultPath $resultPath -Status 'dry_run' -Message 'DryRun completed.' -PlanPath $planPath -Summary $planSummary
    exit 0
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
New-Item -ItemType Directory -Force -Path $epubOutputDir | Out-Null

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$backendExitCode = 0
$allLines = @()
$skipCount = 0
$progressCount = 0
$errorLineCount = 0
$upscaleProgressSummary = [pscustomobject]@{
    total_units = [int]$upscaleEstimate.InputImages
    done_units = 0
    processed_units = 0
    skipped_units = 0
    avg_rate_img_per_sec = 0.0
    elapsed_seconds = 0.0
}
if ($runUpscaleStage) {
    $upscaleStageEstimate = Estimate-UpscaleWork -SourceRoot $sourceDir.FullName -OutputRoot $outputDir -OutputExt (Get-OutputImageExtension -Format $effectiveOutputFormat) -FileNamePattern $outputFilenamePattern
    Write-GuiStageEvent -Stage 'upscale' -Phase 'start' -Data ([pscustomobject]@{
            predicted_process = [int]$upscaleStageEstimate.PredictedProcess
            predicted_skip = [int]$upscaleStageEstimate.PredictedSkip
            total = [int]$upscaleStageEstimate.InputImages
            output_dir = $outputDir
        })
    Write-Host ''
    Write-Host '========== Stage: Upscale =========='
    Write-Host ("Forecast: process={0} skip={1} total={2}" -f $upscaleStageEstimate.PredictedProcess, $upscaleStageEstimate.PredictedSkip, $upscaleStageEstimate.InputImages)
    Write-Host '===================================='
    $upscaleStage = Invoke-UpscaleStage -PythonExe $pythonExe -BackendScript $backendScript -RuntimeSettingsPath $runtimeSettingsPath -LogPath $logPath -TotalUnits ([int]$upscaleStageEstimate.InputImages) -ProgressConfig $progressConfig -LogLevel $LogLevel -NoUpscaleProgress:$NoUpscaleProgress
    $backendExitCode = $upscaleStage.ExitCode
    $allLines = @($upscaleStage.Lines)
    $skipCount = [int]$upscaleStage.SkipCount
    $progressCount = [int]$upscaleStage.ProcessedCount
    $errorLineCount = [int]$upscaleStage.ErrorLineCount
    $upscaleProgressSummary = $upscaleStage.ProgressSummary
    Write-GuiStageEvent -Stage 'upscale' -Phase 'end' -Data ([pscustomobject]@{
            exit_code = [int]$backendExitCode
            processed = [int]$progressCount
            skipped = [int]$skipCount
            done = [int]$upscaleProgressSummary.done_units
            total = [int]$upscaleProgressSummary.total_units
            elapsed_seconds = [double]$upscaleProgressSummary.elapsed_seconds
            avg_rate_img_per_sec = [double]$upscaleProgressSummary.avg_rate_img_per_sec
        })
}
else {
    Write-Info 'Upscale stage skipped.'
    Write-GuiStageEvent -Stage 'upscale' -Phase 'skip' -Data ([pscustomobject]@{ reason = 'disabled' })
}

$epubPlanned = 0
$epubPacked = 0
$epubSkipped = 0
$epubFailed = 0

if ((-not $runUpscaleStage -or $backendExitCode -eq 0) -and $runEpubStage) {
    $epubStageEstimate = Estimate-EpubPackWork -UpscaledRoot $outputDir -PublishedOutputRoot $epubOutputDir -ComicTitle $comicMeta.ComicTitle -TankobonGroupName $GroupTankobon -DefaultGroupName $GroupDefault -VolumeLabel $VolumeChar
    Write-GuiStageEvent -Stage 'epub' -Phase 'start' -Data ([pscustomobject]@{
            predicted_process = [int]$epubStageEstimate.PredictedProcess
            predicted_skip = [int]$epubStageEstimate.PredictedSkip
            total = [int]$epubStageEstimate.Planned
            output_dir = $epubOutputDir
        })
    Write-Host ''
    Write-Host '========== Stage: EPUB Packaging =========='
    Write-Host ("Forecast: process={0} skip={1} total={2}" -f $epubStageEstimate.PredictedProcess, $epubStageEstimate.PredictedSkip, $epubStageEstimate.Planned)
    Write-Host '==========================================='
    Write-StageMessage -Message "EPUB_META: title='$($comicMeta.ComicTitle)' author='$($comicMeta.Author)' source='$($comicMeta.MetadataPath)'" -LogPath $logPath -DebugOnly

    $epubProgressEnabled = $epubStageEstimate.Planned -gt 0
    $epubInteractive = [Environment]::UserInteractive
    $epubNonInteractiveInterval = [int]$progressConfig.NonInteractiveLogIntervalSeconds
    if ($epubNonInteractiveInterval -lt 1) { $epubNonInteractiveInterval = 10 }
    $epubLastTextUpdate = [datetime]::MinValue
    $epubProgressStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $groupNames = @($GroupTankobon, $GroupDefault)
    foreach ($group in $groupNames) {
        $sourceGroupPath = Join-Path $outputDir $group
        if (-not (Test-Path -LiteralPath $sourceGroupPath -PathType Container)) { continue }
        $targetGroupPath = if ($group -eq $GroupTankobon) { $epubOutputDir } else { $sourceGroupPath }
        New-Item -ItemType Directory -Force -Path $targetGroupPath | Out-Null

        $chapterDirs = Get-ChildItem -LiteralPath $sourceGroupPath -Directory | Sort-Object Name
        foreach ($chapterDir in $chapterDirs) {
            if (-not (Test-ChapterHasImages -ChapterPath $chapterDir.FullName)) { continue }

            $epubName = Get-EpubFileName -GroupName $group -ChapterDirName $chapterDir.Name -ComicTitle $comicMeta.ComicTitle -TankobonGroupName $GroupTankobon -VolumeLabel $VolumeChar
            $targetPath = Join-Path $targetGroupPath $epubName
            $epubTitle = [System.IO.Path]::GetFileNameWithoutExtension($epubName)
            $epubPlanned += 1

            if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
                $epubSkipped += 1
                Write-StageMessage -Message "EPUB_SKIP: $targetPath" -LogPath $logPath -DebugOnly
                $epubDone = $epubPacked + $epubSkipped + $epubFailed
                if ($epubProgressEnabled) {
                    $epubSnap = Get-EpubPackProgressSnapshot -DoneUnits $epubDone -TotalUnits $epubStageEstimate.Planned -PackedUnits $epubPacked -SkippedUnits $epubSkipped -FailedUnits $epubFailed -ElapsedSeconds $epubProgressStopwatch.Elapsed.TotalSeconds
                    $epubStatus = ("{0} | {1}" -f $epubTitle, $epubSnap.StatusText)
                    if ($script:PipelineGuiMode) {
                        Write-GuiEvent -Type 'epub_progress' -Data ([pscustomobject]@{
                                percent = [int]$epubSnap.Percent
                                done_units = [int]$epubDone
                                total_units = [int]$epubStageEstimate.Planned
                                packed_units = [int]$epubPacked
                                skipped_units = [int]$epubSkipped
                                failed_units = [int]$epubFailed
                                current_title = $epubTitle
                                status = $epubStatus
                            })
                    }
                    if ($epubInteractive) {
                        Write-Progress -Id 2 -Activity 'Packaging EPUB (KCC)' -Status $epubStatus -PercentComplete $epubSnap.Percent
                    }
                    elseif (((Get-Date) - $epubLastTextUpdate).TotalSeconds -ge $epubNonInteractiveInterval -or $epubDone -eq $epubStageEstimate.Planned) {
                        Write-Host ("EPUB_PROGRESS: {0}% {1}" -f $epubSnap.Percent, $epubStatus)
                        $epubLastTextUpdate = Get-Date
                    }
                }
                continue
            }

            Write-StageMessage -Message "EPUB_PACK: $($chapterDir.FullName) => $targetPath" -LogPath $logPath -DebugOnly
            $kccCode = Invoke-KccPack -KccPath $KccExePath -InputChapterPath $chapterDir.FullName -OutputDirectory $targetGroupPath -Title $epubTitle -Author $comicMeta.Author -TargetEpubPath $targetPath -LogPath $logPath -KccBaseArgs @($pipelineConfig.kcc.EffectiveBaseArgs) -EnableUnicodeStaging ([bool]$pipelineConfig.kcc.UnicodeStaging.Enabled) -UnicodeStagePrefix $pipelineConfig.kcc.UnicodeStaging.StagePrefix -SafeTitleFallback $pipelineConfig.kcc.UnicodeStaging.SafeTitleFallback -SafeAuthorFallback $pipelineConfig.kcc.UnicodeStaging.SafeAuthorFallback -EnableMetadataRewrite ([bool]$pipelineConfig.kcc.MetadataRewrite.Enabled)

            if ($kccCode -ne 0) {
                $epubFailed += 1
                Write-StageMessage -Message "EPUB_FAIL: exit=$kccCode target=$targetPath" -LogPath $logPath -Warning
                $epubDone = $epubPacked + $epubSkipped + $epubFailed
                if ($epubProgressEnabled) {
                    $epubSnap = Get-EpubPackProgressSnapshot -DoneUnits $epubDone -TotalUnits $epubStageEstimate.Planned -PackedUnits $epubPacked -SkippedUnits $epubSkipped -FailedUnits $epubFailed -ElapsedSeconds $epubProgressStopwatch.Elapsed.TotalSeconds
                    $epubStatus = ("{0} | {1}" -f $epubTitle, $epubSnap.StatusText)
                    if ($script:PipelineGuiMode) {
                        Write-GuiEvent -Type 'epub_progress' -Data ([pscustomobject]@{
                                percent = [int]$epubSnap.Percent
                                done_units = [int]$epubDone
                                total_units = [int]$epubStageEstimate.Planned
                                packed_units = [int]$epubPacked
                                skipped_units = [int]$epubSkipped
                                failed_units = [int]$epubFailed
                                current_title = $epubTitle
                                status = $epubStatus
                            })
                    }
                    if ($epubInteractive) {
                        Write-Progress -Id 2 -Activity 'Packaging EPUB (KCC)' -Status $epubStatus -PercentComplete $epubSnap.Percent
                    }
                    elseif (((Get-Date) - $epubLastTextUpdate).TotalSeconds -ge $epubNonInteractiveInterval -or $epubDone -eq $epubStageEstimate.Planned) {
                        Write-Host ("EPUB_PROGRESS: {0}% {1}" -f $epubSnap.Percent, $epubStatus)
                        $epubLastTextUpdate = Get-Date
                    }
                }
                continue
            }

            if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
                $epubPacked += 1
            }
            else {
                $epubFailed += 1
                Write-StageMessage -Message "EPUB_FAIL: kcc succeeded but file missing: $targetPath" -LogPath $logPath -Warning
            }

            $epubDone = $epubPacked + $epubSkipped + $epubFailed
            if ($epubProgressEnabled) {
                $epubSnap = Get-EpubPackProgressSnapshot -DoneUnits $epubDone -TotalUnits $epubStageEstimate.Planned -PackedUnits $epubPacked -SkippedUnits $epubSkipped -FailedUnits $epubFailed -ElapsedSeconds $epubProgressStopwatch.Elapsed.TotalSeconds
                $epubStatus = ("{0} | {1}" -f $epubTitle, $epubSnap.StatusText)
                if ($script:PipelineGuiMode) {
                    Write-GuiEvent -Type 'epub_progress' -Data ([pscustomobject]@{
                            percent = [int]$epubSnap.Percent
                            done_units = [int]$epubDone
                            total_units = [int]$epubStageEstimate.Planned
                            packed_units = [int]$epubPacked
                            skipped_units = [int]$epubSkipped
                            failed_units = [int]$epubFailed
                            current_title = $epubTitle
                            status = $epubStatus
                        })
                }
                if ($epubInteractive) {
                    Write-Progress -Id 2 -Activity 'Packaging EPUB (KCC)' -Status $epubStatus -PercentComplete $epubSnap.Percent
                }
                elseif (((Get-Date) - $epubLastTextUpdate).TotalSeconds -ge $epubNonInteractiveInterval -or $epubDone -eq $epubStageEstimate.Planned) {
                    Write-Host ("EPUB_PROGRESS: {0}% {1}" -f $epubSnap.Percent, $epubStatus)
                    $epubLastTextUpdate = Get-Date
                }
            }
        }
    }
    $epubProgressStopwatch.Stop()
    if ($epubProgressEnabled) {
        $epubDone = $epubPacked + $epubSkipped + $epubFailed
        $epubSnap = Get-EpubPackProgressSnapshot -DoneUnits $epubDone -TotalUnits $epubStageEstimate.Planned -PackedUnits $epubPacked -SkippedUnits $epubSkipped -FailedUnits $epubFailed -ElapsedSeconds $epubProgressStopwatch.Elapsed.TotalSeconds
        $epubStatus = $epubSnap.StatusText
        if ($script:PipelineGuiMode) {
            Write-GuiEvent -Type 'epub_progress' -Data ([pscustomobject]@{
                    percent = [int]$epubSnap.Percent
                    done_units = [int]$epubDone
                    total_units = [int]$epubStageEstimate.Planned
                    packed_units = [int]$epubPacked
                    skipped_units = [int]$epubSkipped
                    failed_units = [int]$epubFailed
                    current_title = $null
                    status = $epubStatus
                })
        }
        if ($epubInteractive) {
            Write-Progress -Id 2 -Activity 'Packaging EPUB (KCC)' -Status $epubStatus -PercentComplete $epubSnap.Percent
            Write-Progress -Id 2 -Activity 'Packaging EPUB (KCC)' -Completed
        }
        else {
            Write-Host ("EPUB_PROGRESS: {0}% {1}" -f $epubSnap.Percent, $epubStatus)
        }
    }
    Write-GuiStageEvent -Stage 'epub' -Phase 'end' -Data ([pscustomobject]@{
            planned = [int]$epubPlanned
            packed = [int]$epubPacked
            skipped = [int]$epubSkipped
            failed = [int]$epubFailed
        })
}
if ($runEpubStage -and $runUpscaleStage -and $backendExitCode -ne 0) {
    Write-Warn ("EPUB packaging stage skipped because upscale failed (BackendExitCode={0})." -f $backendExitCode)
    Write-GuiStageEvent -Stage 'epub' -Phase 'skip' -Data ([pscustomobject]@{ reason = 'upscale_failed'; backend_exit_code = [int]$backendExitCode })
}
if (-not $runEpubStage) {
    Write-GuiStageEvent -Stage 'epub' -Phase 'skip' -Data ([pscustomobject]@{ reason = 'disabled' })
}
$mergedEpubPlanned = 0
$mergedEpubPacked = 0
$mergedEpubSkipped = 0
$mergedEpubFailed = 0
$mergedEpubPath = ''

if ((-not $runUpscaleStage -or $backendExitCode -eq 0) -and $runMergeStage) {
    $defaultGroupPath = Join-Path $outputDir $GroupDefault
    $mergePlanRaw = Get-MergedEpubPlan -SourceRoot $sourceDir.FullName -UpscaledDefaultGroupPath $defaultGroupPath -MergedOutputDefaultGroupPath $epubOutputDefaultGroupPath -ComicTitle $comicMeta.ComicTitle -DefaultGroupName $GroupDefault -MetadataFileName $MetadataJsonName -ChapterMetadataFileName $ChapterMetadataJsonName -TalkLabel $TalkChar
    $mergeOrderAppliedRuntime = Apply-MergeOrderOverrideToPlan -MergePlan $mergePlanRaw -OrderFilePath $MergeOrderFilePath
    $mergePlan = $mergeOrderAppliedRuntime.Plan
    $mergeOrderStateRuntime = $mergeOrderAppliedRuntime.State
    if ($mergeOrderStateRuntime.Exists -and -not $mergeOrderStateRuntime.IsValid) {
        throw ("Merge order file invalid at runtime: {0} ; {1}" -f $MergeOrderFilePath, (($mergeOrderStateRuntime.Errors) -join ' | '))
    }
    foreach ($warnMsg in @($mergeOrderStateRuntime.Warnings)) {
        Write-Warn $warnMsg
    }
    $manifestPath = Join-Path $epubOutputDefaultGroupPath $mergeManifestFileName
    $mergeDecision = if ($mergePlan.Chapters.Count -gt 0) {
        Get-MergedRebuildDecision -ManifestPath $manifestPath -ManifestHash $mergePlan.ManifestHash -TargetPath $mergePlan.TargetPath -MergedFiles $mergePlan.MergedFiles
    }
    else {
        [pscustomobject]@{ NeedRebuild = $false; Reason = 'no-chapters'; OtherMerged = @() }
    }

    Write-Host ''
    Write-Host '========== Stage: Merge EPUB =========='
    Show-MergePreview -MergePlan $mergePlan -Compact:$MergePreviewCompact
    Write-Host ("Forecast: chapters={0} target={1} rebuild={2} reason={3}" -f $mergePlan.Chapters.Count, $mergePlan.TargetPath, $mergeDecision.NeedRebuild, $mergeDecision.Reason)
    Write-Host '======================================='
    Write-GuiStageEvent -Stage 'merge' -Phase 'start' -Data ([pscustomobject]@{
            chapters = [int]$mergePlan.Chapters.Count
            target_path = $mergePlan.TargetPath
            need_rebuild = [bool]$mergeDecision.NeedRebuild
            reason = $mergeDecision.Reason
            order_file = $MergeOrderFilePath
            order_file_exists = [bool]$mergeOrderStateRuntime.Exists
            order_file_applied = [bool]$mergeOrderStateRuntime.Applied
        })
    if ($mergePlan.Chapters.Count -gt 0 -and -not $AutoConfirm) {
        if (-not [Environment]::UserInteractive) {
            Write-GuiEvent -Type 'confirmation_required' -Data ([pscustomobject]@{
                    scope = 'merge_order'
                    interactive_required = $true
                    message = 'Merge order confirmation requires interactive mode or -AutoConfirm.'
                })
            Write-RunArtifacts -ResultPath $resultPath -Status 'needs_confirmation' -Message 'Merge order confirmation requires interactive mode or -AutoConfirm.' -PlanPath $planPath -Summary $planSummary
            throw 'Merge stage confirmation required. Re-run interactively or with -AutoConfirm.'
        }
        Write-GuiEvent -Type 'confirmation_required' -Data ([pscustomobject]@{
                scope = 'merge_order'
                message = 'Merge order shown above. Continue merge stage?'
                default = 'N'
            })
        $mergeConfirm = Read-Host 'Merge order shown above. Continue merge stage? [y/N]'
        Write-GuiEvent -Type 'confirmation_response' -Data ([pscustomobject]@{
                scope = 'merge_order'
                response = $mergeConfirm
                accepted = ($mergeConfirm.ToUpperInvariant() -eq 'Y')
            })
        if ($mergeConfirm.ToUpperInvariant() -ne 'Y') {
            Write-RunArtifacts -ResultPath $resultPath -Status 'aborted_merge_confirm' -Message 'User cancelled at merge order confirmation.' -PlanPath $planPath -Summary $planSummary
            throw 'Execution cancelled: merge order confirmation not approved.'
        }
    }

    if ($mergePlan.Chapters.Count -gt 0) {
        $mergedEpubPlanned = 1
        $mergedEpubPath = $mergePlan.TargetPath

        $decision = $mergeDecision

        if (-not $decision.NeedRebuild) {
            $mergedEpubSkipped += 1
            Write-StageMessage -Message "MERGED_SKIP: target=$($mergePlan.TargetPath) reason=$($decision.Reason)" -LogPath $logPath
        }
        else {
            Write-StageMessage -Message "MERGED_PACK: target=$($mergePlan.TargetPath) reason=$($decision.Reason)" -LogPath $logPath

            $mergeTitle = [System.IO.Path]::GetFileNameWithoutExtension($mergePlan.TargetFileName)
            $mergeAuthor = if ([string]::IsNullOrWhiteSpace($MergedEpubAuthorFallback)) { $comicMeta.Author } else { $MergedEpubAuthorFallback }
            $descLines = @([string]$pipelineConfig.merge.DescriptionHeader)
            foreach ($ch in $mergePlan.Chapters) {
                if ([bool]$pipelineConfig.merge.IncludeOrderInDescription -and $ch.HasOrder) {
                    $descLines += ("{0} (order {1})" -f $ch.ChapterName, (Format-OrderValue -Value ([double]$ch.Order)))
                }
                else {
                    $descLines += $ch.ChapterName
                }
            }

            $mergePlanJsonPath = Join-Path $logsDir ("merge_plan_{0}.json" -f $timestamp)
            $mergePlanObject = [pscustomobject]@{
                output_epub_path = $mergePlan.TargetPath
                title = $mergeTitle
                author = $mergeAuthor
                language = [string]$pipelineConfig.merge.Language
                description = ($descLines -join "`n")
                contributor = [string]$pipelineConfig.merge.MetadataContributor
                chapters = @($mergePlan.Chapters | ForEach-Object {
                        if ($_.HasOrder) {
                            [pscustomobject]@{ chapter_name = $_.ChapterName; order = [double]$_.Order; epub_path = $_.EpubPath }
                        }
                        else {
                            [pscustomobject]@{ chapter_name = $_.ChapterName; order = $null; epub_path = $_.EpubPath }
                        }
                    })
            }
            Write-JsonUtf8NoBom -Object $mergePlanObject -Path $mergePlanJsonPath

            $mergeCode = Invoke-MergedEpubPack -PythonExe $pythonExe -MergeScriptPath $mergeScriptPath -PlanPath $mergePlanJsonPath -LogPath $logPath
            if ($mergeCode -ne 0) {
                $mergedEpubFailed += 1
                Write-StageMessage -Message "MERGED_FAIL: exit=$mergeCode target=$($mergePlan.TargetPath)" -LogPath $logPath -Warning
            }
            elseif (-not (Test-Path -LiteralPath $mergePlan.TargetPath -PathType Leaf)) {
                $mergedEpubFailed += 1
                Write-StageMessage -Message "MERGED_FAIL: merge succeeded but target missing: $($mergePlan.TargetPath)" -LogPath $logPath -Warning
            }
            else {
                $deleteFailed = $false
                foreach ($oldFile in $decision.OtherMerged) {
                    try {
                        Remove-Item -LiteralPath $oldFile.FullName -Force
                        Write-StageMessage -Message "MERGED_DELETE_OLD: $($oldFile.FullName)" -LogPath $logPath -DebugOnly
                    }
                    catch {
                        $deleteFailed = $true
                        Write-StageMessage -Message "MERGED_DELETE_FAIL: $($oldFile.FullName) ; $($_.Exception.Message)" -LogPath $logPath -Warning
                    }
                }

                if ($deleteFailed) {
                    $mergedEpubFailed += 1
                }
                else {
                    $mergedEpubPacked += 1
                    $manifestObject = [pscustomobject]@{
                        ManifestHash = $mergePlan.ManifestHash
                        MergeOrderFilePath = $MergeOrderFilePath
                        MergeOrderSignature = if ($mergeOrderStateRuntime.Exists -and $mergeOrderStateRuntime.IsValid) { $mergeOrderStateRuntime.Signature } else { 'AUTO' }
                        MergedFileName = $mergePlan.TargetFileName
                        MergedFilePath = $mergePlan.TargetPath
                        ChapterCount = $mergePlan.Chapters.Count
                        GeneratedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                        Chapters = @($mergePlan.Chapters | ForEach-Object {
                                [pscustomobject]@{
                                    chapter_name = $_.ChapterName
                                    order = if ($_.HasOrder) { [double]$_.Order } else { $null }
                                    epub_path = $_.EpubPath
                                }
                            })
                    }
                    Write-JsonUtf8NoBom -Object $manifestObject -Path $manifestPath
                }
            }
        }
    }
    else {
        Write-StageMessage -Message 'MERGED_SKIP: no chapter EPUB candidates in default group.' -LogPath $logPath
    }
    Write-GuiStageEvent -Stage 'merge' -Phase 'end' -Data ([pscustomobject]@{
            planned = [int]$mergedEpubPlanned
            packed = [int]$mergedEpubPacked
            skipped = [int]$mergedEpubSkipped
            failed = [int]$mergedEpubFailed
            target_path = $mergedEpubPath
        })
}
if ($runMergeStage -and $runUpscaleStage -and $backendExitCode -ne 0) {
    Write-Warn ("Merged EPUB stage skipped because upscale failed (BackendExitCode={0})." -f $backendExitCode)
    Write-GuiStageEvent -Stage 'merge' -Phase 'skip' -Data ([pscustomobject]@{ reason = 'upscale_failed'; backend_exit_code = [int]$backendExitCode })
}
if (-not $runMergeStage) {
    Write-GuiStageEvent -Stage 'merge' -Phase 'skip' -Data ([pscustomobject]@{ reason = 'disabled' })
}

$stopwatch.Stop()

Write-Host ''
Write-Host '========== Summary =========='
Write-Host ("BackendExitCode:  {0}" -f $backendExitCode)
Write-Host ("Elapsed:          {0:N2}s" -f $stopwatch.Elapsed.TotalSeconds)
Write-Host ("Skipped images:   {0}" -f $skipCount)
Write-Host ("Progress lines:   {0}" -f $progressCount)
Write-Host ("Error-like lines: {0}" -f $errorLineCount)
Write-Host ("UpscaleDone/Total:{0}/{1}" -f $upscaleProgressSummary.done_units, $upscaleProgressSummary.total_units)
Write-Host ("UpscaleAvgRate:   {0:N2} img/s" -f $upscaleProgressSummary.avg_rate_img_per_sec)
Write-Host ("EpubPlanned:      {0}" -f $epubPlanned)
Write-Host ("EpubPacked:       {0}" -f $epubPacked)
Write-Host ("EpubSkipped:      {0}" -f $epubSkipped)
Write-Host ("EpubFailed:       {0}" -f $epubFailed)
Write-Host ("MergedEpubPlanned:{0}" -f $mergedEpubPlanned)
Write-Host ("MergedEpubPacked: {0}" -f $mergedEpubPacked)
Write-Host ("MergedEpubSkipped:{0}" -f $mergedEpubSkipped)
Write-Host ("MergedEpubFailed: {0}" -f $mergedEpubFailed)
Write-Host ("MergedEpubPath:   {0}" -f $mergedEpubPath)
Write-Host ("Log file:         {0}" -f $logPath)
Write-Host '============================='

$runResultSummary = [pscustomobject]@{
    Plan = $planSummary
    Runtime = [pscustomobject]@{
        BackendExitCode = [int]$backendExitCode
        ElapsedSeconds = [double]$stopwatch.Elapsed.TotalSeconds
        SkipCount = [int]$skipCount
        ProgressCount = [int]$progressCount
        ErrorLineCount = [int]$errorLineCount
        EpubPlanned = [int]$epubPlanned
        EpubPacked = [int]$epubPacked
        EpubSkipped = [int]$epubSkipped
        EpubFailed = [int]$epubFailed
        MergedEpubPlanned = [int]$mergedEpubPlanned
        MergedEpubPacked = [int]$mergedEpubPacked
        MergedEpubSkipped = [int]$mergedEpubSkipped
        MergedEpubFailed = [int]$mergedEpubFailed
        MergedEpubPath = $mergedEpubPath
        MergeOrderFilePath = $MergeOrderFilePath
        UpscaleProgressSummary = $upscaleProgressSummary
    }
}

if ($backendExitCode -ne 0) {
    Write-RunArtifacts -ResultPath $resultPath -Status 'backend_failed' -Message ("Backend exited with non-zero code: {0}" -f $backendExitCode) -PlanPath $planPath -Summary $runResultSummary
    throw "Backend exited with non-zero code: $backendExitCode. See log: $logPath"
}

if ($epubFailed -gt 0) {
    Write-RunArtifacts -ResultPath $resultPath -Status 'epub_failed' -Message ("EPUB packaging failed for {0} item(s)." -f $epubFailed) -PlanPath $planPath -Summary $runResultSummary
    throw "EPUB packaging failed for $epubFailed item(s). See log: $logPath"
}

if ($mergedEpubFailed -gt 0) {
    Write-RunArtifacts -ResultPath $resultPath -Status 'merge_failed' -Message ("Merged EPUB stage failed for {0} item(s)." -f $mergedEpubFailed) -PlanPath $planPath -Summary $runResultSummary
    throw "Merged EPUB stage failed for $mergedEpubFailed item(s). See log: $logPath"
}

Write-RunArtifacts -ResultPath $resultPath -Status 'success' -Message 'Execution completed successfully.' -PlanPath $planPath -Summary $runResultSummary

