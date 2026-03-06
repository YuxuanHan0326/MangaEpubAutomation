using System.Collections.ObjectModel;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MangaEpubAutomation.Gui;
using MangaEpubAutomation.Gui.Models;
using MangaEpubAutomation.Gui.Services;

namespace MangaEpubAutomation.Gui.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ProcessRunner _processRunner = new();
    private readonly string _repoRoot;
    private readonly StringBuilder _logBuilder = new();
    private CancellationTokenSource? _runningCts;
    private static LocalizationManager Loc => LocalizationManager.Instance;

    public MainViewModel()
    {
        _repoRoot = ResolveRepoRoot();
        ScriptPath = Path.Combine(_repoRoot, "Invoke-MangaEpubAutomation.ps1");
        DepsConfigPath = Path.Combine(_repoRoot, "manga_epub_automation.deps.json");
        ConfigPath = Path.Combine(_repoRoot, "manga_epub_automation.config.json");
        UpscaleFactor = 2;
        OutputFormat = "webp";
        LossyQuality = 80;
        LogLevel = "info";
        RunUpscale = true;
        RunEpubPackaging = true;
        RunMergedEpub = true;
        DepsProgressEnabled = true;
        DepsRefreshSeconds = 1;
        DepsEtaMinSamples = 8;
        DepsNoninteractiveLogIntervalSeconds = 10;
        CfgSourceDirSuffixToSkip = "-upscaled";
        CfgOutputDirSuffix = "-upscaled";
        CfgEpubOutputDirSuffix = "-output";
        CfgOutputFilenamePattern = "%filename%-upscaled";
        CfgOverwriteExistingFiles = false;
        CfgUpscaleScaleFactor = 2;
        CfgOutputFormat = "webp";
        CfgLossyCompressionQuality = 80;
        CfgKccDeviceProfile = "KS";
        CfgKccOutputFormat = "EPUB";
        CfgKccNoKepub = true;
        CfgKccDisableProcessing = true;
        CfgKccForceColor = true;
        CfgKccImageOutputMode = 0;
        CfgKccSplitterMode = -1;
        CfgKccCroppingMode = -1;
        CfgKccInterPanelCropMode = -1;
        CfgKccBatchSplitMode = -1;
        CfgKccMetadataTitleMode = -1;
        CfgKccGamma = "Auto";
        CfgMergeLanguage = "zh";
        CfgMergeDescriptionHeader = "选集 包含:";
        CfgMergeIncludeOrderInDescription = true;
        CfgMergeMetadataContributor = "manga-epub-automation-merge";

        BrowseTitleRootCommand = new RelayCommand(BrowseTitleRoot);
        BrowseScriptPathCommand = new RelayCommand(BrowseScriptPath);
        BrowseDepsConfigPathCommand = new RelayCommand(BrowseDepsConfigPath);
        BrowsePipelineConfigPathCommand = new RelayCommand(BrowsePipelineConfigPath);
        BrowseMergeOrderFilePathCommand = new RelayCommand(BrowseMergeOrderFilePath);
        BrowseDepsPythonExeCommand = new RelayCommand(BrowseDepsPythonExe);
        BrowseDepsBackendScriptCommand = new RelayCommand(BrowseDepsBackendScript);
        BrowseDepsModelsDirCommand = new RelayCommand(BrowseDepsModelsDir);
        BrowseDepsKccExeCommand = new RelayCommand(BrowseDepsKccExe);

        UseFullPipelinePresetCommand = new RelayCommand(() => SetStagePreset(true, true, true));
        UseUpscaleOnlyPresetCommand = new RelayCommand(() => SetStagePreset(true, false, false));
        UseEpubOnlyPresetCommand = new RelayCommand(() => SetStagePreset(false, true, false));
        UseMergeOnlyPresetCommand = new RelayCommand(() => SetStagePreset(false, false, true));
        UseFrontHalfPresetCommand = new RelayCommand(() => SetStagePreset(true, true, false));
        UseBackHalfPresetCommand = new RelayCommand(() => SetStagePreset(false, true, true));

        GeneratePlanCommand = new AsyncRelayCommand(GeneratePlanAsync, () => !IsRunning);
        ExecuteRunCommand = new AsyncRelayCommand(ExecutePipelineAsync, () => !IsRunning);
        CancelRunCommand = new RelayCommand(() => _runningCts?.Cancel(), () => IsRunning);
        LoadLatestArtifactsCommand = new RelayCommand(LoadLatestArtifacts);

        LoadDependenciesCommand = new RelayCommand(LoadDependenciesFromFile);
        SaveDependenciesCommand = new RelayCommand(SaveDependenciesFromForm);
        SaveDependenciesRawJsonCommand = new RelayCommand(SaveDependenciesRawJson);
        InitDependenciesFromTemplateCommand = new RelayCommand(InitDepsFromTemplate);
        OpenMangaJaNaiRepoCommand = new RelayCommand(() => OpenUrl("https://github.com/the-database/MangaJaNaiConverterGui"));
        OpenKccRepoCommand = new RelayCommand(() => OpenUrl("https://github.com/ciromattia/kcc"));
        OpenCopyMangaDownloaderRepoCommand = new RelayCommand(() => OpenUrl("https://github.com/misaka10843/copymanga-downloader"));

        LoadPipelineConfigCommand = new RelayCommand(LoadPipelineConfigFromFile);
        SavePipelineConfigCommand = new RelayCommand(SavePipelineConfigFromForm);
        SavePipelineConfigRawJsonCommand = new RelayCommand(SavePipelineConfigRawJson);
        InitPipelineConfigCommand = new RelayCommand(InitConfigViaScript);

        LoadMergeOrderCommand = new RelayCommand(LoadMergeOrderFile);
        SaveMergeOrderCommand = new RelayCommand(SaveMergeOrderFile);
        GenerateMergeOrderTemplateCommand = new AsyncRelayCommand(GenerateMergeOrderTemplateAsync, () => !IsRunning);
        AddMergeOrderEntryCommand = new RelayCommand(AddMergeOrderEntry);
        RemoveMergeOrderEntryCommand = new RelayCommand(RemoveMergeOrderEntry, () => SelectedMergeOrderEntry is not null);
        MoveMergeOrderUpCommand = new RelayCommand(MoveMergeOrderEntryUp, () => SelectedMergeOrderEntry is not null);
        MoveMergeOrderDownCommand = new RelayCommand(MoveMergeOrderEntryDown, () => SelectedMergeOrderEntry is not null);

        ClearLogsCommand = new RelayCommand(ClearLogs);
        LoadLatestPlanCommand = new RelayCommand(LoadLatestPlanPreview);
        LoadLatestResultCommand = new RelayCommand(LoadLatestResultPreview);
        Loc.PropertyChanged += OnLocalizationChanged;

        if (File.Exists(DepsConfigPath)) LoadDependenciesFromFile();
        if (File.Exists(ConfigPath)) LoadPipelineConfigFromFile();
        RunStatusText = L("Status.Idle");
        PreflightSummaryText = L("Status.NoPreflight");
    }

    public IReadOnlyList<int> UpscaleFactorOptions { get; } = new[] { 1, 2, 3, 4 };
    public IReadOnlyList<string> OutputFormatOptions { get; } = new[] { "webp", "png", "jpeg", "avif" };
    public IReadOnlyList<string> LogLevelOptions { get; } = new[] { "info", "debug" };

    public ObservableCollection<PipelineIssue> PreflightIssues { get; } = new();
    public ObservableCollection<MergeChapterItem> MergePreviewChapters { get; } = new();
    public ObservableCollection<string> MergeOrderChapters { get; } = new();

    [ObservableProperty] private string scriptPath = string.Empty;
    [ObservableProperty] private string titleRoot = string.Empty;
    [ObservableProperty] private string sourceDirName = string.Empty;
    [ObservableProperty] private string depsConfigPath = string.Empty;
    [ObservableProperty] private string configPath = string.Empty;
    [ObservableProperty] private string mergeOrderFilePath = string.Empty;
    [ObservableProperty] private int upscaleFactor;
    [ObservableProperty] private string outputFormat = string.Empty;
    [ObservableProperty] private int lossyQuality;
    [ObservableProperty] private string logLevel = string.Empty;
    [ObservableProperty] private bool dryRun;
    [ObservableProperty] private bool noUpscaleProgress;
    [ObservableProperty] private bool failOnPreflightWarnings;
    [ObservableProperty] private bool mergePreviewCompact;
    [ObservableProperty] private bool runUpscale;
    [ObservableProperty] private bool runEpubPackaging;
    [ObservableProperty] private bool runMergedEpub;
    [ObservableProperty] private bool isRunning;
    [ObservableProperty] private string runStatusText = string.Empty;
    [ObservableProperty] private double upscalePercent;
    [ObservableProperty] private string upscaleStatusText = "-";
    [ObservableProperty] private double epubPercent;
    [ObservableProperty] private string epubStatusText = "-";
    [ObservableProperty] private double mergePercent;
    [ObservableProperty] private string mergeStatusText = "-";
    [ObservableProperty] private bool mergeProgressIndeterminate;
    [ObservableProperty] private string preflightSummaryText = "No preflight result yet.";
    [ObservableProperty] private string mergeTargetPath = string.Empty;
    [ObservableProperty] private string lastPlanPath = string.Empty;
    [ObservableProperty] private string lastResultPath = string.Empty;
    [ObservableProperty] private string depsJsonText = string.Empty;
    [ObservableProperty] private string pipelineConfigJsonText = string.Empty;
    [ObservableProperty] private string depsPythonExe = string.Empty;
    [ObservableProperty] private string depsBackendScript = string.Empty;
    [ObservableProperty] private string depsModelsDir = string.Empty;
    [ObservableProperty] private string depsKccExe = string.Empty;
    [ObservableProperty] private bool depsProgressEnabled;
    [ObservableProperty] private int depsRefreshSeconds;
    [ObservableProperty] private int depsEtaMinSamples;
    [ObservableProperty] private int depsNoninteractiveLogIntervalSeconds;
    [ObservableProperty] private string cfgSourceDirSuffixToSkip = string.Empty;
    [ObservableProperty] private string cfgOutputDirSuffix = string.Empty;
    [ObservableProperty] private string cfgEpubOutputDirSuffix = string.Empty;
    [ObservableProperty] private string cfgOutputFilenamePattern = string.Empty;
    [ObservableProperty] private bool cfgOverwriteExistingFiles;
    [ObservableProperty] private int cfgUpscaleScaleFactor;
    [ObservableProperty] private string cfgOutputFormat = string.Empty;
    [ObservableProperty] private int cfgLossyCompressionQuality;
    [ObservableProperty] private string cfgKccDeviceProfile = string.Empty;
    [ObservableProperty] private string cfgKccOutputFormat = string.Empty;
    [ObservableProperty] private bool cfgKccNoKepub;
    [ObservableProperty] private bool cfgKccDisableProcessing;
    [ObservableProperty] private bool cfgKccForceColor;
    [ObservableProperty] private bool cfgKccMangaStyle;
    [ObservableProperty] private bool cfgKccWebtoonMode;
    [ObservableProperty] private bool cfgKccTwoPanelView;
    [ObservableProperty] private bool cfgKccHighQualityMagnification;
    [ObservableProperty] private string cfgKccTargetSizeMb = string.Empty;
    [ObservableProperty] private bool cfgKccLegacyPdfExtract;
    [ObservableProperty] private bool cfgKccUpscaleSmallImages;
    [ObservableProperty] private bool cfgKccStretchToResolution;
    [ObservableProperty] private int cfgKccSplitterMode;
    [ObservableProperty] private string cfgKccGamma = string.Empty;
    [ObservableProperty] private int cfgKccCroppingMode;
    [ObservableProperty] private string cfgKccCroppingPower = string.Empty;
    [ObservableProperty] private string cfgKccPreserveMarginPercent = string.Empty;
    [ObservableProperty] private string cfgKccCroppingMinimumRatio = string.Empty;
    [ObservableProperty] private int cfgKccInterPanelCropMode;
    [ObservableProperty] private bool cfgKccForceBlackBorders;
    [ObservableProperty] private bool cfgKccForceWhiteBorders;
    [ObservableProperty] private int cfgKccImageOutputMode;
    [ObservableProperty] private string cfgKccJpegQuality = string.Empty;
    [ObservableProperty] private bool cfgKccMaximizeStrips;
    [ObservableProperty] private int cfgKccBatchSplitMode;
    [ObservableProperty] private int cfgKccMetadataTitleMode;
    [ObservableProperty] private bool cfgKccSpreadShift;
    [ObservableProperty] private bool cfgKccNoRotateSpreads;
    [ObservableProperty] private bool cfgKccRotateFirstSpread;
    [ObservableProperty] private bool cfgKccAutoLevel;
    [ObservableProperty] private bool cfgKccDisableAutoContrast;
    [ObservableProperty] private bool cfgKccColorAutoContrast;
    [ObservableProperty] private bool cfgKccFileFusion;
    [ObservableProperty] private bool cfgKccEraseRainbow;
    [ObservableProperty] private bool cfgKccDeleteSourceAfterPack;
    [ObservableProperty] private string cfgKccCustomWidth = string.Empty;
    [ObservableProperty] private string cfgKccCustomHeight = string.Empty;
    [ObservableProperty] private string cfgMergeLanguage = string.Empty;
    [ObservableProperty] private string cfgMergeDescriptionHeader = string.Empty;
    [ObservableProperty] private bool cfgMergeIncludeOrderInDescription;
    [ObservableProperty] private string cfgMergeMetadataContributor = string.Empty;
    [ObservableProperty] private string newMergeOrderEntry = string.Empty;
    [ObservableProperty] private string? selectedMergeOrderEntry;
    [ObservableProperty] private string liveLogText = string.Empty;
    [ObservableProperty] private string planJsonPreview = string.Empty;
    [ObservableProperty] private string resultJsonPreview = string.Empty;

    public IRelayCommand BrowseTitleRootCommand { get; }
    public IRelayCommand BrowseScriptPathCommand { get; }
    public IRelayCommand BrowseDepsConfigPathCommand { get; }
    public IRelayCommand BrowsePipelineConfigPathCommand { get; }
    public IRelayCommand BrowseMergeOrderFilePathCommand { get; }
    public IRelayCommand BrowseDepsPythonExeCommand { get; }
    public IRelayCommand BrowseDepsBackendScriptCommand { get; }
    public IRelayCommand BrowseDepsModelsDirCommand { get; }
    public IRelayCommand BrowseDepsKccExeCommand { get; }
    public IRelayCommand UseFullPipelinePresetCommand { get; }
    public IRelayCommand UseUpscaleOnlyPresetCommand { get; }
    public IRelayCommand UseEpubOnlyPresetCommand { get; }
    public IRelayCommand UseMergeOnlyPresetCommand { get; }
    public IRelayCommand UseFrontHalfPresetCommand { get; }
    public IRelayCommand UseBackHalfPresetCommand { get; }
    public IAsyncRelayCommand GeneratePlanCommand { get; }
    public IAsyncRelayCommand ExecuteRunCommand { get; }
    public IRelayCommand CancelRunCommand { get; }
    public IRelayCommand LoadLatestArtifactsCommand { get; }
    public IRelayCommand LoadDependenciesCommand { get; }
    public IRelayCommand SaveDependenciesCommand { get; }
    public IRelayCommand SaveDependenciesRawJsonCommand { get; }
    public IRelayCommand InitDependenciesFromTemplateCommand { get; }
    public IRelayCommand OpenMangaJaNaiRepoCommand { get; }
    public IRelayCommand OpenKccRepoCommand { get; }
    public IRelayCommand OpenCopyMangaDownloaderRepoCommand { get; }
    public IRelayCommand LoadPipelineConfigCommand { get; }
    public IRelayCommand SavePipelineConfigCommand { get; }
    public IRelayCommand SavePipelineConfigRawJsonCommand { get; }
    public IRelayCommand InitPipelineConfigCommand { get; }
    public IRelayCommand LoadMergeOrderCommand { get; }
    public IRelayCommand SaveMergeOrderCommand { get; }
    public IAsyncRelayCommand GenerateMergeOrderTemplateCommand { get; }
    public IRelayCommand AddMergeOrderEntryCommand { get; }
    public IRelayCommand RemoveMergeOrderEntryCommand { get; }
    public IRelayCommand MoveMergeOrderUpCommand { get; }
    public IRelayCommand MoveMergeOrderDownCommand { get; }
    public IRelayCommand ClearLogsCommand { get; }
    public IRelayCommand LoadLatestPlanCommand { get; }
    public IRelayCommand LoadLatestResultCommand { get; }

    partial void OnIsRunningChanged(bool value)
    {
        GeneratePlanCommand.NotifyCanExecuteChanged();
        ExecuteRunCommand.NotifyCanExecuteChanged();
        CancelRunCommand.NotifyCanExecuteChanged();
        GenerateMergeOrderTemplateCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedMergeOrderEntryChanged(string? value)
    {
        RemoveMergeOrderEntryCommand.NotifyCanExecuteChanged();
        MoveMergeOrderUpCommand.NotifyCanExecuteChanged();
        MoveMergeOrderDownCommand.NotifyCanExecuteChanged();
    }

    private static string ResolveRepoRoot()
    {
        static string? Walk(string start)
        {
            var current = new DirectoryInfo(start);
            while (current is not null)
            {
                if (File.Exists(Path.Combine(current.FullName, "Invoke-MangaEpubAutomation.ps1")))
                    return current.FullName;
                current = current.Parent;
            }
            return null;
        }

        return Walk(AppContext.BaseDirectory) ?? Walk(Environment.CurrentDirectory) ?? Environment.CurrentDirectory;
    }

    private void SetStagePreset(bool upscale, bool epub, bool merge)
    {
        RunUpscale = upscale;
        RunEpubPackaging = epub;
        RunMergedEpub = merge;
    }

    private void BrowseTitleRoot()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = L("Msg.SelectTitleRoot"),
            SelectedPath = Directory.Exists(TitleRoot) ? TitleRoot : _repoRoot
        };
        if (dialog.ShowDialog() == DialogResult.OK) TitleRoot = dialog.SelectedPath;
    }

    private void BrowseScriptPath()
    {
        var path = BrowseFilePath("PowerShell Script (*.ps1)|*.ps1", ScriptPath, true);
        if (!string.IsNullOrWhiteSpace(path)) ScriptPath = path;
    }

    private void BrowseDepsConfigPath()
    {
        var path = BrowseFilePath("JSON (*.json)|*.json|All files (*.*)|*.*", DepsConfigPath, true);
        if (!string.IsNullOrWhiteSpace(path))
        {
            DepsConfigPath = path;
            LoadJsonEditor(DepsConfigPath, v => DepsJsonText = v);
        }
    }

    private void BrowsePipelineConfigPath()
    {
        var path = BrowseFilePath("JSON (*.json)|*.json|All files (*.*)|*.*", ConfigPath, true);
        if (!string.IsNullOrWhiteSpace(path))
        {
            ConfigPath = path;
            LoadJsonEditor(ConfigPath, v => PipelineConfigJsonText = v);
        }
    }

    private void BrowseMergeOrderFilePath()
    {
        var path = BrowseFilePath("JSON (*.json)|*.json|All files (*.*)|*.*", MergeOrderFilePath, false);
        if (!string.IsNullOrWhiteSpace(path)) MergeOrderFilePath = path;
    }

    private void BrowseDepsPythonExe()
    {
        var path = BrowseFilePath("Python executable (python.exe)|python.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*", DepsPythonExe, true);
        if (!string.IsNullOrWhiteSpace(path)) DepsPythonExe = path;
    }

    private void BrowseDepsBackendScript()
    {
        var path = BrowseFilePath("Python script (*.py)|*.py|All files (*.*)|*.*", DepsBackendScript, true);
        if (!string.IsNullOrWhiteSpace(path)) DepsBackendScript = path;
    }

    private void BrowseDepsModelsDir()
    {
        var path = BrowseFolderPath(DepsModelsDir, L("Msg.SelectModelsDir"));
        if (!string.IsNullOrWhiteSpace(path)) DepsModelsDir = path;
    }

    private void BrowseDepsKccExe()
    {
        var path = BrowseFilePath("Executable (*.exe)|*.exe|All files (*.*)|*.*", DepsKccExe, true);
        if (!string.IsNullOrWhiteSpace(path)) DepsKccExe = path;
    }

    private string? BrowseFilePath(string filter, string currentValue, bool checkExists)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = filter,
            InitialDirectory = GetInitialDirectory(currentValue, _repoRoot),
            CheckFileExists = checkExists
        };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static string? BrowseFolderPath(string currentValue, string description)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = description,
            SelectedPath = GetInitialDirectory(currentValue)
        };
        return dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedPath : null;
    }

    private static string GetInitialDirectory(params string[] values)
    {
        foreach (var value in values)
        {
            if (Directory.Exists(value)) return value;
            if (!string.IsNullOrWhiteSpace(value))
            {
                var parent = Path.GetDirectoryName(value);
                if (!string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent)) return parent;
            }
        }
        return Environment.CurrentDirectory;
    }

    private async Task GeneratePlanAsync()
    {
        if (!ValidateRunInputs(out var error))
        {
            ShowError(error);
            return;
        }
        await RunPipelineProcessAsync(planOnly: true).ConfigureAwait(false);
    }

    private async Task ExecutePipelineAsync()
    {
        if (!ValidateRunInputs(out var error))
        {
            ShowError(error);
            return;
        }

        var planRun = await RunPipelineProcessAsync(planOnly: true).ConfigureAwait(false);
        if (planRun is null || planRun.WasCanceled) return;
        if (GetPreflightCount("ERROR") > 0)
        {
            ShowError(L("Msg.PreflightHasErrors"));
            return;
        }
        if (GetPreflightCount("WARN") > 0)
        {
            var ok = AskYesNo(BuildWarningMessage(), L("Msg.PreflightWarningsTitle"));
            if (!ok) return;
        }

        if (!RunMergedEpub)
        {
            await RunPipelineProcessAsync(planOnly: false).ConfigureAwait(false);
            return;
        }

        var hasPreMergeStages = RunUpscale || RunEpubPackaging;
        if (hasPreMergeStages)
        {
            var preMergeRun = await RunPipelineProcessAsync(
                planOnly: false,
                runUpscaleOverride: RunUpscale,
                runEpubOverride: RunEpubPackaging,
                runMergeOverride: false,
                resetProgress: true).ConfigureAwait(false);
            if (preMergeRun is null || preMergeRun.WasCanceled || preMergeRun.ExitCode != 0) return;

            var mergePlanRefresh = await RunPipelineProcessAsync(
                planOnly: true,
                runUpscaleOverride: false,
                runEpubOverride: false,
                runMergeOverride: true).ConfigureAwait(false);
            if (mergePlanRefresh is null || mergePlanRefresh.WasCanceled) return;
            if (GetPreflightCount("ERROR") > 0)
            {
                ShowError(L("Msg.PreflightHasErrors"));
                return;
            }
        }

        if (MergePreviewChapters.Count > 0)
        {
            var ok = AskYesNo(LF("Msg.MergeConfirmFmt", MergePreviewChapters.Count), L("Msg.MergeConfirmTitle"));
            if (!ok) return;
        }

        await RunPipelineProcessAsync(
            planOnly: false,
            runUpscaleOverride: false,
            runEpubOverride: false,
            runMergeOverride: true,
            resetProgress: false).ConfigureAwait(false);
    }

    private async Task<ProcessRunResult?> RunPipelineProcessAsync(
        bool planOnly,
        bool? runUpscaleOverride = null,
        bool? runEpubOverride = null,
        bool? runMergeOverride = null,
        bool resetProgress = true)
    {
        if (IsRunning) return null;

        var effectiveRunUpscale = runUpscaleOverride ?? RunUpscale;
        var effectiveRunEpub = runEpubOverride ?? RunEpubPackaging;
        var effectiveRunMerge = runMergeOverride ?? RunMergedEpub;

        RunOnUi(() =>
        {
            IsRunning = true;
            RunStatusText = planOnly ? L("Status.RunningPlan") : L("Status.RunningPipeline");
        });
        _runningCts = new CancellationTokenSource();
        if (!planOnly && resetProgress)
        {
            RunOnUi(() =>
            {
                UpscalePercent = 0;
                EpubPercent = 0;
                MergePercent = 0;
                UpscaleStatusText = "-";
                EpubStatusText = "-";
                MergeStatusText = "-";
                MergeProgressIndeterminate = false;
            });
        }

        var args = BuildPipelineArguments(planOnly, effectiveRunUpscale, effectiveRunEpub, effectiveRunMerge);
        AppendLog($"Run: {string.Join(" ", args)}");

        ProcessRunResult result;
        try
        {
            result = await _processRunner.RunPowerShellFileAsync(
                ScriptPath,
                args,
                line => RunOnUi(() => AppendLog(line)),
                line => RunOnUi(() => AppendLog("[stderr] " + line)),
                ev => RunOnUi(() => HandlePipelineEvent(ev)),
                _runningCts.Token).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            RunOnUi(() => ShowError(LF("Msg.LaunchFailedFmt", ex.Message)));
            result = null!;
        }

        RunOnUi(() =>
        {
            if (result is not null)
            {
                RunStatusText = result.WasCanceled ? L("Status.Canceled") : LF("Status.ExitFmt", result.ExitCode);
                AppendLog(LF("Msg.ProcessExitFmt", result.ExitCode));
                LoadLatestArtifacts();
            }
        });

        RunOnUi(() => IsRunning = false);
        _runningCts?.Dispose();
        _runningCts = null;
        return result;
    }

    private List<string> BuildPipelineArguments(bool planOnly, bool runUpscale, bool runEpubPackaging, bool runMergedEpub)
    {
        var args = new List<string>
        {
            "-GuiMode",
            "-AutoConfirm",
            "-TitleRoot", TitleRoot,
            "-UpscaleFactor", UpscaleFactor.ToString(),
            "-OutputFormat", OutputFormat,
            "-LossyQuality", LossyQuality.ToString(),
            "-LogLevel", LogLevel
        };
        if (!string.IsNullOrWhiteSpace(SourceDirName)) { args.Add("-SourceDirName"); args.Add(SourceDirName); }
        if (!string.IsNullOrWhiteSpace(DepsConfigPath)) { args.Add("-DepsConfigPath"); args.Add(DepsConfigPath); }
        if (!string.IsNullOrWhiteSpace(ConfigPath)) { args.Add("-ConfigPath"); args.Add(ConfigPath); }
        if (!string.IsNullOrWhiteSpace(MergeOrderFilePath)) { args.Add("-MergeOrderFilePath"); args.Add(MergeOrderFilePath); }
        if (!runUpscale) args.Add("-SkipUpscale");
        if (!runEpubPackaging) args.Add("-SkipEpubPackaging");
        if (!runMergedEpub) args.Add("-SkipMergedEpub");
        if (DryRun) args.Add("-DryRun");
        if (NoUpscaleProgress) args.Add("-NoUpscaleProgress");
        if (FailOnPreflightWarnings) args.Add("-FailOnPreflightWarnings");
        if (MergePreviewCompact) args.Add("-MergePreviewCompact");
        if (planOnly) args.Add("-PlanOnly");
        return args;
    }

    private bool ValidateRunInputs(out string error)
    {
        if (string.IsNullOrWhiteSpace(ScriptPath) || !File.Exists(ScriptPath))
        {
            error = L("Msg.Validate.ScriptPath");
            return false;
        }
        if (string.IsNullOrWhiteSpace(TitleRoot) || !Directory.Exists(TitleRoot))
        {
            error = L("Msg.Validate.TitleRoot");
            return false;
        }
        if (!RunUpscale && !RunEpubPackaging && !RunMergedEpub)
        {
            error = L("Msg.Validate.StageRequired");
            return false;
        }
        if (RunUpscale && RunMergedEpub && !RunEpubPackaging)
        {
            error = L("Msg.Validate.StageCombination");
            return false;
        }
        if (!TryParseJsonFile(DepsConfigPath, out error))
        {
            error = LF("Msg.Validate.DepsJsonFmt", error);
            return false;
        }
        if (!TryParseJsonFile(ConfigPath, out error))
        {
            error = LF("Msg.Validate.ConfigJsonFmt", error);
            return false;
        }
        if (!string.IsNullOrWhiteSpace(MergeOrderFilePath) && File.Exists(MergeOrderFilePath)
            && !TryReadMergeOrder(MergeOrderFilePath, out _, out error))
        {
            error = LF("Msg.Validate.MergeOrderJsonFmt", error);
            return false;
        }
        error = string.Empty;
        return true;
    }

    private static bool TryParseJsonFile(string path, out string error)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            error = "file not found";
            return false;
        }

        try
        {
            _ = JsonNode.Parse(File.ReadAllText(path, Encoding.UTF8));
            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool ValidateJsonText(string json, out string error)
    {
        try
        {
            _ = JsonNode.Parse(json);
            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private void HandlePipelineEvent(PipelineEventEnvelope ev)
    {
        switch (ev.Type)
        {
            case "stage":
                ParseStageEvent(ev.Data);
                break;
            case "plan_ready":
                LastPlanPath = GetString(ev.Data, "plan_path");
                break;
            case "preflight_summary":
                ParsePreflight(ev.Data);
                break;
            case "merge_preview":
                ParseMergePreview(ev.Data);
                break;
            case "upscale_progress":
                UpscalePercent = GetDouble(ev.Data, "percent");
                UpscaleStatusText = GetString(ev.Data, "status");
                break;
            case "epub_progress":
                EpubPercent = GetDouble(ev.Data, "percent");
                EpubStatusText = GetString(ev.Data, "status");
                break;
            case "run_result":
                LastResultPath = GetString(ev.Data, "result_path");
                break;
            default:
                AppendLog(LF("Msg.EventFmt", ev.Type, ev.RawJson));
                break;
        }
    }

    private void ParseStageEvent(JsonElement data)
    {
        var stage = GetString(data, "stage").Trim().ToLowerInvariant();
        var phase = GetString(data, "phase").Trim().ToLowerInvariant();
        var payload = TryGetProperty(data, "data", out var nested) ? nested : default;

        if (!string.Equals(stage, "merge", StringComparison.Ordinal))
        {
            return;
        }

        switch (phase)
        {
            case "start":
            {
                var chapters = GetInt(payload, "chapters");
                var needRebuild = GetBool(payload, "need_rebuild");
                var reason = GetString(payload, "reason");
                MergePercent = 0;
                MergeProgressIndeterminate = chapters > 0 && needRebuild;

                if (chapters <= 0)
                {
                    MergeStatusText = L("Status.MergeNoChapters");
                }
                else if (needRebuild)
                {
                    MergeStatusText = LF("Status.MergeRunningFmt", chapters);
                }
                else
                {
                    MergePercent = 100;
                    MergeProgressIndeterminate = false;
                    MergeStatusText = LF("Status.MergeSkippedFmt", string.IsNullOrWhiteSpace(reason) ? "up-to-date" : reason);
                }
                break;
            }
            case "end":
            {
                MergePercent = 100;
                MergeProgressIndeterminate = false;
                var packed = GetInt(payload, "packed");
                var skipped = GetInt(payload, "skipped");
                var failed = GetInt(payload, "failed");
                MergeStatusText = LF("Status.MergeDoneFmt", packed, skipped, failed);
                break;
            }
            case "skip":
            {
                MergePercent = 100;
                MergeProgressIndeterminate = false;
                var reason = GetString(payload, "reason");
                MergeStatusText = LF("Status.MergeSkippedFmt", string.IsNullOrWhiteSpace(reason) ? "skipped" : reason);
                break;
            }
        }
    }

    private void ParsePreflight(JsonElement data)
    {
        var errors = GetInt(data, "errors");
        var warnings = GetInt(data, "warnings");
        var infos = GetInt(data, "infos");
        PreflightSummaryText = LF("Run.PreflightSummaryFmt", errors, warnings, infos);
        PreflightIssues.Clear();
        if (TryGetProperty(data, "issues", out var issues) && issues.ValueKind == JsonValueKind.Array)
        {
            foreach (var it in issues.EnumerateArray())
            {
                var severity = GetString(it, "Severity");
                if (string.IsNullOrWhiteSpace(severity)) severity = GetString(it, "severity");
                var code = GetString(it, "Code");
                if (string.IsNullOrWhiteSpace(code)) code = GetString(it, "code");
                var message = GetString(it, "Message");
                if (string.IsNullOrWhiteSpace(message)) message = GetString(it, "message");
                PreflightIssues.Add(new PipelineIssue
                {
                    Severity = NormalizeSeverity(severity),
                    Code = code,
                    Message = message
                });
            }
        }
        RefreshPreflightSummaryFromIssues(errors, warnings, infos);
    }

    private void ParseMergePreview(JsonElement data)
    {
        MergeTargetPath = GetString(data, "target_path");
        MergePreviewChapters.Clear();
        if (TryGetProperty(data, "chapters", out var chapters) && chapters.ValueKind == JsonValueKind.Array)
        {
            foreach (var ch in chapters.EnumerateArray())
            {
                MergePreviewChapters.Add(new MergeChapterItem
                {
                    Index = GetInt(ch, "Index"),
                    Chapter = GetString(ch, "Chapter"),
                    Order = GetString(ch, "Order"),
                    EpubPath = GetString(ch, "EpubPath")
                });
            }
        }
    }

    private static bool TryGetProperty(JsonElement element, string key, out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var p in element.EnumerateObject())
            {
                if (string.Equals(p.Name, key, StringComparison.OrdinalIgnoreCase))
                {
                    value = p.Value;
                    return true;
                }
            }
        }
        value = default;
        return false;
    }

    private static string GetString(JsonElement element, string key) => TryGetProperty(element, key, out var value) ? value.ToString() : string.Empty;
    private static int GetInt(JsonElement element, string key) => TryGetProperty(element, key, out var value) && value.TryGetInt32(out var n) ? n : 0;
    private static double GetDouble(JsonElement element, string key) => TryGetProperty(element, key, out var value) && value.TryGetDouble(out var n) ? n : 0;
    private static bool GetBool(JsonElement element, string key)
    {
        if (!TryGetProperty(element, key, out var value))
        {
            return false;
        }

        if (value.ValueKind == JsonValueKind.True) return true;
        if (value.ValueKind == JsonValueKind.False) return false;
        if (value.ValueKind == JsonValueKind.String)
        {
            var text = value.GetString();
            return string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    private int GetPreflightCount(string severity)
    {
        var expected = NormalizeSeverity(severity);
        return PreflightIssues.Count(x => string.Equals(NormalizeSeverity(x.Severity), expected, StringComparison.OrdinalIgnoreCase));
    }

    private string BuildWarningMessage()
    {
        var lines = PreflightIssues
            .Where(x => string.Equals(NormalizeSeverity(x.Severity), "WARN", StringComparison.OrdinalIgnoreCase))
            .Select(x => $"- [{x.Code}] {x.Message}")
            .ToList();

        if (lines.Count == 0) return L("Msg.PreflightWarningsContinue");
        return L("Msg.PreflightWarningsHeader") + "\n\n" + string.Join("\n", lines) + "\n\n" + L("Msg.PreflightWarningsContinue");
    }

    private void RefreshPreflightSummaryFromIssues(int? errorCount = null, int? warnCount = null, int? infoCount = null)
    {
        var errors = errorCount ?? GetPreflightCount("ERROR");
        var warnings = warnCount ?? GetPreflightCount("WARN");
        var infos = infoCount ?? GetPreflightCount("INFO");
        PreflightSummaryText = LF("Run.PreflightSummaryFmt", errors, warnings, infos);
    }

    private static string NormalizeSeverity(string? severity)
    {
        if (string.IsNullOrWhiteSpace(severity))
        {
            return "INFO";
        }

        var normalized = severity.Trim().ToUpperInvariant();
        return normalized switch
        {
            "WARNING" => "WARN",
            "INFORMATION" => "INFO",
            _ => normalized
        };
    }

    private static bool AskYesNo(string message, string title)
    {
        if (TryShowNativeTaskDialog(
                caption: title,
                heading: title,
                text: message,
                icon: TaskDialogIcon.Warning,
                buttons: [TaskDialogButton.Yes, TaskDialogButton.No],
                defaultButton: TaskDialogButton.No,
                out var result))
        {
            return result == TaskDialogButton.Yes;
        }

        return System.Windows.MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes;
    }

    private static void ShowError(string message)
    {
        if (TryShowNativeTaskDialog(
                caption: L("App.GuiTitle"),
                heading: L("App.GuiTitle"),
                text: message,
                icon: TaskDialogIcon.Error,
                buttons: [TaskDialogButton.OK],
                defaultButton: TaskDialogButton.OK,
                out _))
        {
            return;
        }

        System.Windows.MessageBox.Show(message, L("App.GuiTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static bool TryShowNativeTaskDialog(
        string caption,
        string heading,
        string text,
        TaskDialogIcon icon,
        IReadOnlyList<TaskDialogButton> buttons,
        TaskDialogButton defaultButton,
        out TaskDialogButton result)
    {
        result = TaskDialogButton.Cancel;
        try
        {
            var page = new TaskDialogPage
            {
                Caption = caption,
                Heading = heading,
                Text = text,
                Icon = icon,
                AllowCancel = true
            };
            foreach (var button in buttons)
            {
                page.Buttons.Add(button);
            }
            page.DefaultButton = defaultButton;
            result = TaskDialog.ShowDialog(page);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void AppendLog(string line)
    {
        _logBuilder.AppendLine(line);
        LiveLogText = _logBuilder.ToString();
        CaptureRuntimeWarning(line);
    }

    private void CaptureRuntimeWarning(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return;
        }

        var trimmed = line.Trim();
        if (!trimmed.StartsWith("[WARN]", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var payload = trimmed[6..].Trim();
        var code = "RUNTIME_WARN";
        var message = payload;
        var match = Regex.Match(payload, @"^\[(?<code>[^\]]+)\]\s*(?<msg>.+)$");
        if (match.Success)
        {
            code = match.Groups["code"].Value.Trim();
            message = match.Groups["msg"].Value.Trim();
        }

        var exists = PreflightIssues.Any(x =>
            string.Equals(NormalizeSeverity(x.Severity), "WARN", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.Code, code, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.Message, message, StringComparison.Ordinal));
        if (!exists)
        {
            PreflightIssues.Add(new PipelineIssue
            {
                Severity = "WARN",
                Code = code,
                Message = message
            });
            RefreshPreflightSummaryFromIssues();
        }
    }

    private void ClearLogs()
    {
        _logBuilder.Clear();
        LiveLogText = string.Empty;
    }

    private void LoadLatestArtifacts()
    {
        LoadLatestPlanPreview();
        LoadLatestResultPreview();
    }

    private void LoadLatestPlanPreview()
    {
        var path = Path.Combine(_repoRoot, "logs", "latest_run_plan.json");
        if (File.Exists(path))
        {
            LastPlanPath = path;
            var text = File.ReadAllText(path, Encoding.UTF8);
            PlanJsonPreview = text;
            TryPopulatePreflightFromPlan(text);
        }
    }

    private void LoadLatestResultPreview()
    {
        var path = Path.Combine(_repoRoot, "logs", "latest_run_result.json");
        if (File.Exists(path))
        {
            LastResultPath = path;
            ResultJsonPreview = File.ReadAllText(path, Encoding.UTF8);
        }
    }

    private void TryPopulatePreflightFromPlan(string planJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(planJson);
            var root = doc.RootElement;
            if (!TryGetProperty(root, "Preflight", out var preflight) || preflight.ValueKind != JsonValueKind.Object)
            {
                return;
            }

            var errors = GetInt(preflight, "ErrorCount");
            var warnings = GetInt(preflight, "WarnCount");
            var infos = GetInt(preflight, "InfoCount");

            PreflightIssues.Clear();
            if (TryGetProperty(preflight, "Issues", out var issues) && issues.ValueKind == JsonValueKind.Array)
            {
                foreach (var it in issues.EnumerateArray())
                {
                    var severity = GetString(it, "Severity");
                    var code = GetString(it, "Code");
                    var message = GetString(it, "Message");
                    PreflightIssues.Add(new PipelineIssue
                    {
                        Severity = NormalizeSeverity(severity),
                        Code = code,
                        Message = message
                    });
                }
            }

            RefreshPreflightSummaryFromIssues(errors, warnings, infos);
        }
        catch
        {
            // Keep previous UI state when latest plan JSON can't be parsed.
        }
    }

    private void LoadJsonEditor(string path, Action<string> assign)
    {
        if (!File.Exists(path))
        {
            ShowError(LF("Msg.FileNotFoundFmt", path));
            return;
        }
        assign(File.ReadAllText(path, Encoding.UTF8));
        AppendLog(LF("Msg.LoadedFileFmt", path));
    }

    private void SaveJsonEditor(string path, string jsonText)
    {
        if (!ValidateJsonText(jsonText, out var error))
        {
            ShowError(LF("Msg.JsonInvalidFmt", error));
            return;
        }
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
        File.WriteAllText(path, JsonNode.Parse(jsonText)!.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
        AppendLog(LF("Msg.SavedFileFmt", path));
    }

    private void LoadDependenciesFromFile()
    {
        if (!File.Exists(DepsConfigPath))
        {
            ShowError(LF("Msg.FileNotFoundFmt", DepsConfigPath));
            return;
        }

        var text = File.ReadAllText(DepsConfigPath, Encoding.UTF8);
        DepsJsonText = text;

        try
        {
            var root = JsonNode.Parse(text)?.AsObject() ?? new JsonObject();
            var paths = root["paths"]?.AsObject() ?? new JsonObject();
            var progress = root["progress"]?.AsObject() ?? new JsonObject();

            DepsPythonExe = GetJsonString(paths, "python_exe");
            DepsBackendScript = GetJsonString(paths, "backend_script");
            DepsModelsDir = GetJsonString(paths, "models_dir");
            DepsKccExe = GetJsonString(paths, "kcc_exe");

            DepsProgressEnabled = GetJsonBool(progress, "enabled", true);
            DepsRefreshSeconds = GetJsonInt(progress, "refresh_seconds", 1);
            DepsEtaMinSamples = GetJsonInt(progress, "eta_min_samples", 8);
            DepsNoninteractiveLogIntervalSeconds = GetJsonInt(progress, "noninteractive_log_interval_seconds", 10);

            AppendLog("Loaded dependencies: " + DepsConfigPath);
        }
        catch (Exception ex)
        {
            ShowError(LF("Msg.DepsParseFailedFmt", ex.Message));
        }
    }

    private void SaveDependenciesFromForm()
    {
        var root = new JsonObject
        {
            ["version"] = 1,
            ["paths"] = new JsonObject
            {
                ["python_exe"] = DepsPythonExe,
                ["backend_script"] = DepsBackendScript,
                ["models_dir"] = DepsModelsDir,
                ["kcc_exe"] = DepsKccExe
            },
            ["progress"] = new JsonObject
            {
                ["enabled"] = DepsProgressEnabled,
                ["refresh_seconds"] = DepsRefreshSeconds,
                ["eta_min_samples"] = DepsEtaMinSamples,
                ["noninteractive_log_interval_seconds"] = DepsNoninteractiveLogIntervalSeconds
            }
        };

        var json = root.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
        SaveJsonEditor(DepsConfigPath, json);
        DepsJsonText = json;
    }

    private void SaveDependenciesRawJson()
    {
        SaveJsonEditor(DepsConfigPath, DepsJsonText);
        LoadDependenciesFromFile();
    }

    private void LoadPipelineConfigFromFile()
    {
        if (!File.Exists(ConfigPath))
        {
            ShowError(LF("Msg.FileNotFoundFmt", ConfigPath));
            return;
        }

        var text = File.ReadAllText(ConfigPath, Encoding.UTF8);
        PipelineConfigJsonText = text;

        try
        {
            var root = JsonNode.Parse(text)?.AsObject() ?? new JsonObject();
            var pipeline = root["pipeline"]?.AsObject() ?? new JsonObject();
            var manga = root["manga"]?.AsObject() ?? new JsonObject();
            var kcc = root["kcc"]?.AsObject() ?? new JsonObject();
            var cli = kcc["CliOptions"]?.AsObject() ?? new JsonObject();
            var merge = root["merge"]?.AsObject() ?? new JsonObject();

            CfgSourceDirSuffixToSkip = GetJsonString(pipeline, "SourceDirSuffixToSkip", "-upscaled");
            CfgOutputDirSuffix = GetJsonString(pipeline, "OutputDirSuffix", "-upscaled");
            CfgEpubOutputDirSuffix = GetJsonString(pipeline, "EpubOutputDirSuffix", "-output");
            CfgOutputFilenamePattern = GetJsonString(pipeline, "OutputFilenamePattern", "%filename%-upscaled");

            CfgOverwriteExistingFiles = GetJsonBool(manga, "OverwriteExistingFiles", false);
            CfgUpscaleScaleFactor = GetJsonInt(manga, "UpscaleScaleFactor", 2);
            CfgOutputFormat = GetJsonString(manga, "OutputFormat", "webp");
            CfgLossyCompressionQuality = GetJsonInt(manga, "LossyCompressionQuality", 80);

            CfgKccDeviceProfile = GetJsonString(cli, "DeviceProfile", "KS");
            CfgKccOutputFormat = GetJsonString(cli, "OutputFormat", "EPUB");
            CfgKccNoKepub = GetJsonBool(cli, "NoKepub", true);
            CfgKccDisableProcessing = GetJsonBool(cli, "DisableProcessing", true);
            CfgKccForceColor = GetJsonBool(cli, "ForceColor", true);
            CfgKccMangaStyle = GetJsonBool(cli, "MangaStyle", false);
            CfgKccWebtoonMode = GetJsonBool(cli, "WebtoonMode", false);
            CfgKccTwoPanelView = GetJsonBool(cli, "TwoPanelView", false);
            CfgKccHighQualityMagnification = GetJsonBool(cli, "HighQualityMagnification", false);
            CfgKccTargetSizeMb = GetJsonValueAsText(cli, "TargetSizeMB", string.Empty);
            CfgKccLegacyPdfExtract = GetJsonBool(cli, "LegacyPdfExtract", false);
            CfgKccUpscaleSmallImages = GetJsonBool(cli, "UpscaleSmallImages", false);
            CfgKccStretchToResolution = GetJsonBool(cli, "StretchToResolution", false);
            CfgKccSplitterMode = GetJsonNullableInt(cli, "SplitterMode") ?? -1;
            CfgKccGamma = GetJsonValueAsText(cli, "Gamma", "Auto");
            CfgKccCroppingMode = GetJsonNullableInt(cli, "CroppingMode") ?? -1;
            CfgKccCroppingPower = GetJsonValueAsText(cli, "CroppingPower", string.Empty);
            CfgKccPreserveMarginPercent = GetJsonValueAsText(cli, "PreserveMarginPercent", string.Empty);
            CfgKccCroppingMinimumRatio = GetJsonValueAsText(cli, "CroppingMinimumRatio", string.Empty);
            CfgKccInterPanelCropMode = GetJsonNullableInt(cli, "InterPanelCropMode") ?? -1;
            CfgKccForceBlackBorders = GetJsonBool(cli, "ForceBlackBorders", false);
            CfgKccForceWhiteBorders = GetJsonBool(cli, "ForceWhiteBorders", false);
            var forcePng = GetJsonBool(cli, "ForcePng", false);
            var mozJpeg = GetJsonBool(cli, "MozJpeg", false);
            CfgKccImageOutputMode = forcePng ? 1 : mozJpeg ? 2 : 0;
            CfgKccJpegQuality = GetJsonValueAsText(cli, "JpegQuality", string.Empty);
            CfgKccMaximizeStrips = GetJsonBool(cli, "MaximizeStrips", false);
            CfgKccBatchSplitMode = GetJsonNullableInt(cli, "BatchSplitMode") ?? -1;
            CfgKccMetadataTitleMode = GetJsonNullableInt(cli, "MetadataTitleMode") ?? -1;
            CfgKccSpreadShift = GetJsonBool(cli, "SpreadShift", false);
            CfgKccNoRotateSpreads = GetJsonBool(cli, "NoRotateSpreads", false);
            CfgKccRotateFirstSpread = GetJsonBool(cli, "RotateFirstSpread", false);
            CfgKccAutoLevel = GetJsonBool(cli, "AutoLevel", false);
            CfgKccDisableAutoContrast = GetJsonBool(cli, "DisableAutoContrast", false);
            CfgKccColorAutoContrast = GetJsonBool(cli, "ColorAutoContrast", false);
            CfgKccFileFusion = GetJsonBool(cli, "FileFusion", false);
            CfgKccEraseRainbow = GetJsonBool(cli, "EraseRainbow", false);
            CfgKccDeleteSourceAfterPack = GetJsonBool(cli, "DeleteSourceAfterPack", false);
            CfgKccCustomWidth = GetJsonValueAsText(cli, "CustomWidth", string.Empty);
            CfgKccCustomHeight = GetJsonValueAsText(cli, "CustomHeight", string.Empty);

            CfgMergeLanguage = GetJsonString(merge, "Language", "zh");
            CfgMergeDescriptionHeader = GetJsonString(merge, "DescriptionHeader", "选集 包含:");
            CfgMergeIncludeOrderInDescription = GetJsonBool(merge, "IncludeOrderInDescription", true);
            CfgMergeMetadataContributor = GetJsonString(merge, "MetadataContributor", "manga-epub-automation-merge");

            AppendLog("Loaded config: " + ConfigPath);
        }
        catch (Exception ex)
        {
            ShowError(LF("Msg.ConfigParseFailedFmt", ex.Message));
        }
    }

    private void SavePipelineConfigFromForm()
    {
        var targetSizeMb = ParseNullableInt(CfgKccTargetSizeMb);
        var splitterMode = ParseNullableMode(CfgKccSplitterMode);
        var croppingMode = ParseNullableMode(CfgKccCroppingMode);
        var interPanelCropMode = ParseNullableMode(CfgKccInterPanelCropMode);
        var batchSplitMode = ParseNullableMode(CfgKccBatchSplitMode);
        var metadataTitleMode = ParseNullableMode(CfgKccMetadataTitleMode);
        var gammaValue = ParseNullableDoubleFromText(CfgKccGamma, "Auto");
        var croppingPowerValue = ParseNullableDouble(CfgKccCroppingPower);
        var preserveMarginPercentValue = ParseNullableDouble(CfgKccPreserveMarginPercent);
        var croppingMinimumRatioValue = ParseNullableDouble(CfgKccCroppingMinimumRatio);
        var jpegQuality = ParseNullableInt(CfgKccJpegQuality);
        var customWidth = ParseNullableInt(CfgKccCustomWidth);
        var customHeight = ParseNullableInt(CfgKccCustomHeight);
        var forcePng = CfgKccImageOutputMode == 1;
        var mozJpeg = CfgKccImageOutputMode == 2;

        var root = new JsonObject
        {
            ["version"] = 2,
            ["pipeline"] = new JsonObject
            {
                ["SourceDirSuffixToSkip"] = CfgSourceDirSuffixToSkip,
                ["OutputDirSuffix"] = CfgOutputDirSuffix,
                ["EpubOutputDirSuffix"] = CfgEpubOutputDirSuffix,
                ["OutputFilenamePattern"] = CfgOutputFilenamePattern
            },
            ["manga"] = new JsonObject
            {
                ["SelectedTabIndex"] = 1,
                ["OverwriteExistingFiles"] = CfgOverwriteExistingFiles,
                ["ModeScaleSelected"] = true,
                ["ModeWidthSelected"] = false,
                ["ModeHeightSelected"] = false,
                ["ModeFitToDisplaySelected"] = false,
                ["UpscaleScaleFactor"] = CfgUpscaleScaleFactor,
                ["OutputFormat"] = CfgOutputFormat,
                ["LossyCompressionQuality"] = CfgLossyCompressionQuality,
                ["WorkflowOverrides"] = new JsonObject()
            },
            ["kcc"] = new JsonObject
            {
                ["CliOptions"] = new JsonObject
                {
                    ["DeviceProfile"] = CfgKccDeviceProfile,
                    ["OutputFormat"] = CfgKccOutputFormat,
                    ["NoKepub"] = CfgKccNoKepub,
                    ["DisableProcessing"] = CfgKccDisableProcessing,
                    ["ForceColor"] = CfgKccForceColor,
                    ["MangaStyle"] = CfgKccMangaStyle,
                    ["HighQualityMagnification"] = CfgKccHighQualityMagnification,
                    ["TwoPanelView"] = CfgKccTwoPanelView,
                    ["WebtoonMode"] = CfgKccWebtoonMode,
                    ["TargetSizeMB"] = targetSizeMb,
                    ["LegacyPdfExtract"] = CfgKccLegacyPdfExtract,
                    ["UpscaleSmallImages"] = CfgKccUpscaleSmallImages,
                    ["StretchToResolution"] = CfgKccStretchToResolution,
                    ["SplitterMode"] = splitterMode,
                    ["Gamma"] = gammaValue,
                    ["CroppingMode"] = croppingMode,
                    ["CroppingPower"] = croppingPowerValue,
                    ["PreserveMarginPercent"] = preserveMarginPercentValue,
                    ["CroppingMinimumRatio"] = croppingMinimumRatioValue,
                    ["InterPanelCropMode"] = interPanelCropMode,
                    ["ForceBlackBorders"] = CfgKccForceBlackBorders,
                    ["ForceWhiteBorders"] = CfgKccForceWhiteBorders,
                    ["ForcePng"] = forcePng,
                    ["MozJpeg"] = mozJpeg,
                    ["JpegQuality"] = jpegQuality,
                    ["MaximizeStrips"] = CfgKccMaximizeStrips,
                    ["BatchSplitMode"] = batchSplitMode,
                    ["MetadataTitleMode"] = metadataTitleMode,
                    ["SpreadShift"] = CfgKccSpreadShift,
                    ["NoRotateSpreads"] = CfgKccNoRotateSpreads,
                    ["RotateFirstSpread"] = CfgKccRotateFirstSpread,
                    ["AutoLevel"] = CfgKccAutoLevel,
                    ["DisableAutoContrast"] = CfgKccDisableAutoContrast,
                    ["ColorAutoContrast"] = CfgKccColorAutoContrast,
                    ["FileFusion"] = CfgKccFileFusion,
                    ["EraseRainbow"] = CfgKccEraseRainbow,
                    ["DeleteSourceAfterPack"] = CfgKccDeleteSourceAfterPack,
                    ["CustomWidth"] = customWidth,
                    ["CustomHeight"] = customHeight,
                    ["AdditionalArgs"] = new JsonArray()
                },
                ["UnicodeStaging"] = new JsonObject
                {
                    ["Enabled"] = true,
                    ["StagePrefix"] = "manga_epub_automation_kcc_stage_",
                    ["SafeTitleFallback"] = "manga_epub_automation_book",
                    ["SafeAuthorFallback"] = "pipeline"
                },
                ["MetadataRewrite"] = new JsonObject
                {
                    ["Enabled"] = true
                },
                ["BaseArgs"] = new JsonArray()
            },
            ["merge"] = new JsonObject
            {
                ["Language"] = CfgMergeLanguage,
                ["DescriptionHeader"] = CfgMergeDescriptionHeader,
                ["IncludeOrderInDescription"] = CfgMergeIncludeOrderInDescription,
                ["MetadataContributor"] = CfgMergeMetadataContributor
            }
        };

        var json = root.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
        SaveJsonEditor(ConfigPath, json);
        PipelineConfigJsonText = json;
    }

    private void SavePipelineConfigRawJson()
    {
        SaveJsonEditor(ConfigPath, PipelineConfigJsonText);
        LoadPipelineConfigFromFile();
    }

    private static string GetJsonString(JsonObject obj, string key, string defaultValue = "")
    {
        var node = obj[key];
        return node?.GetValue<string>() ?? defaultValue;
    }

    private static int GetJsonInt(JsonObject obj, string key, int defaultValue = 0)
    {
        var node = obj[key];
        if (node is null) return defaultValue;
        try
        {
            return node.GetValue<int>();
        }
        catch
        {
        }
        return defaultValue;
    }

    private static int? GetJsonNullableInt(JsonObject obj, string key)
    {
        var node = obj[key];
        if (node is null || node.ToJsonString().Equals("null", StringComparison.OrdinalIgnoreCase)) return null;

        try { return node.GetValue<int>(); } catch { }
        try
        {
            if (int.TryParse(node.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)) return parsed;
        }
        catch { }
        return null;
    }

    private static string GetJsonValueAsText(JsonObject obj, string key, string defaultValue)
    {
        var node = obj[key];
        if (node is null || node.ToJsonString().Equals("null", StringComparison.OrdinalIgnoreCase)) return defaultValue;

        try { return node.GetValue<string>(); } catch { }
        try { return node.GetValue<int>().ToString(CultureInfo.InvariantCulture); } catch { }
        try { return node.GetValue<double>().ToString(CultureInfo.InvariantCulture); } catch { }
        try { return node.GetValue<decimal>().ToString(CultureInfo.InvariantCulture); } catch { }
        try { return node.GetValue<bool>() ? "true" : "false"; } catch { }
        return defaultValue;
    }

    private static bool GetJsonBool(JsonObject obj, string key, bool defaultValue = false)
    {
        var node = obj[key];
        if (node is null) return defaultValue;
        try
        {
            return node.GetValue<bool>();
        }
        catch
        {
        }
        return defaultValue;
    }

    private static int? ParseNullableInt(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        return int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    private static int? ParseNullableMode(int modeValue)
    {
        return modeValue >= 0 ? modeValue : null;
    }

    private static double? ParseNullableDouble(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        return double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    private static double? ParseNullableDoubleFromText(string? text, params string[] nullTokens)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var normalized = text.Trim();
        foreach (var token in nullTokens)
        {
            if (normalized.Equals(token, StringComparison.OrdinalIgnoreCase)) return null;
        }
        return ParseNullableDouble(normalized);
    }

    private void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            ShowError(LF("Msg.OpenUrlFailedFmt", ex.Message));
        }
    }

    private void InitDepsFromTemplate()
    {
        var template = Path.Combine(_repoRoot, "manga_epub_automation.deps.template.json");
        if (!File.Exists(template))
        {
            ShowError(LF("Msg.TemplateNotFoundFmt", template));
            return;
        }
        if (File.Exists(DepsConfigPath) && !AskYesNo("Deps file exists. Overwrite from template?", "Init Deps")) return;
        Directory.CreateDirectory(Path.GetDirectoryName(DepsConfigPath)!);
        File.Copy(template, DepsConfigPath, true);
        LoadDependenciesFromFile();
    }

    private void InitConfigViaScript()
    {
        if (!File.Exists(ScriptPath))
        {
            ShowError(L("Msg.ScriptNotFound"));
            return;
        }
        _ = Task.Run(async () =>
        {
            var args = new List<string> { "-InitConfig", "-ConfigPath", ConfigPath };
            var result = await _processRunner.RunPowerShellFileAsync(
                ScriptPath, args,
                line => RunOnUi(() => AppendLog(line)),
                line => RunOnUi(() => AppendLog("[stderr] " + line)),
                _ => { },
                CancellationToken.None);
            RunOnUi(() =>
            {
                AppendLog(LF("Msg.InitConfigExitFmt", result.ExitCode));
                if (File.Exists(ConfigPath)) LoadPipelineConfigFromFile();
            });
        });
    }

    private void LoadMergeOrderFile()
    {
        if (!TryReadMergeOrder(MergeOrderFilePath, out var chapters, out var error))
        {
            ShowError(error);
            return;
        }
        MergeOrderChapters.Clear();
        foreach (var ch in chapters) MergeOrderChapters.Add(ch);
    }

    private void SaveMergeOrderFile()
    {
        if (string.IsNullOrWhiteSpace(MergeOrderFilePath))
        {
            ShowError(L("Msg.MergeOrderPathEmpty"));
            return;
        }
        var root = new JsonObject
        {
            ["version"] = 1,
            ["chapters"] = new JsonArray(MergeOrderChapters.Select(x => (JsonNode?)x).ToArray())
        };
        Directory.CreateDirectory(Path.GetDirectoryName(MergeOrderFilePath)!);
        File.WriteAllText(MergeOrderFilePath, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
        AppendLog(LF("Msg.MergeOrderSavedFmt", MergeOrderFilePath));
    }

    private async Task GenerateMergeOrderTemplateAsync()
    {
        if (MergePreviewChapters.Count == 0)
        {
            var plan = await RunPipelineProcessAsync(planOnly: true).ConfigureAwait(false);
            if (plan is null || plan.WasCanceled) return;
        }
        if (MergePreviewChapters.Count == 0)
        {
            ShowError(L("Msg.MergePreviewEmpty"));
            return;
        }
        if (string.IsNullOrWhiteSpace(MergeOrderFilePath))
        {
            MergeOrderFilePath = Path.Combine(TitleRoot, ".manga_epub_automation_merge_order.json");
        }
        MergeOrderChapters.Clear();
        foreach (var ch in MergePreviewChapters) MergeOrderChapters.Add(ch.Chapter);
        SaveMergeOrderFile();
    }

    private void AddMergeOrderEntry()
    {
        var value = NewMergeOrderEntry?.Trim();
        if (!string.IsNullOrWhiteSpace(value))
        {
            MergeOrderChapters.Add(value);
            NewMergeOrderEntry = string.Empty;
        }
    }

    private void RemoveMergeOrderEntry()
    {
        if (SelectedMergeOrderEntry is not null) MergeOrderChapters.Remove(SelectedMergeOrderEntry);
    }

    private void MoveMergeOrderEntryUp()
    {
        if (SelectedMergeOrderEntry is null) return;
        var index = MergeOrderChapters.IndexOf(SelectedMergeOrderEntry);
        if (index > 0) MergeOrderChapters.Move(index, index - 1);
    }

    private void MoveMergeOrderEntryDown()
    {
        if (SelectedMergeOrderEntry is null) return;
        var index = MergeOrderChapters.IndexOf(SelectedMergeOrderEntry);
        if (index >= 0 && index < MergeOrderChapters.Count - 1) MergeOrderChapters.Move(index, index + 1);
    }

    private static bool TryReadMergeOrder(string path, out List<string> chapters, out string error)
    {
        chapters = new List<string>();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            error = LF("Msg.FileNotFoundFmt", path);
            return false;
        }
        try
        {
            var root = JsonNode.Parse(File.ReadAllText(path, Encoding.UTF8))?.AsObject();
            var array = root?["chapters"]?.AsArray();
            if (array is null)
            {
                error = L("Msg.MergeOrderMissingChapters");
                return false;
            }
            foreach (var item in array)
            {
                if (item is null) continue;
                var value = item.GetValue<string>().Trim();
                if (!string.IsNullOrWhiteSpace(value)) chapters.Add(value);
            }
            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static void RunOnUi(Action action)
    {
        if (System.Windows.Application.Current.Dispatcher.CheckAccess()) action();
        else System.Windows.Application.Current.Dispatcher.Invoke(action);
    }

    private void OnLocalizationChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.Equals(e.PropertyName, "Item[]", StringComparison.Ordinal))
        {
            return;
        }

        RunOnUi(() =>
        {
            if (!IsRunning)
            {
                RunStatusText = L("Status.Idle");
            }
            if (PreflightIssues.Count == 0)
            {
                PreflightSummaryText = L("Status.NoPreflight");
            }
        });
    }

    private static string L(string key) => Loc.Get(key);

    private static string LF(string key, params object[] args) => Loc.Format(key, args);
}
