using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MangaEpubAutomation.Gui.Models;

namespace MangaEpubAutomation.Gui.ViewModels;

public partial class MainViewModel
{
    public ObservableCollection<string> EpubInputFiles { get; } = new();
    public ObservableCollection<PipelineIssue> EpubPreflightIssues { get; } = new();

    [ObservableProperty] private int epubInputMode; // 0 = files, 1 = directory
    [ObservableProperty] private string epubInputDirectory = string.Empty;
    [ObservableProperty] private string? selectedEpubInputFile;
    [ObservableProperty] private string epubOutputDirectory = string.Empty;
    [ObservableProperty] private string epubOutputSuffix = "-upscaled";
    [ObservableProperty] private bool epubOverwriteOutputEpub;
    [ObservableProperty] private string epubRunStatusText = string.Empty;
    [ObservableProperty] private double epubFilePercent;
    [ObservableProperty] private string epubFileStatusText = "-";
    [ObservableProperty] private string epubPreflightSummaryText = string.Empty;
    [ObservableProperty] private string epubLiveLogText = string.Empty;
    [ObservableProperty] private string epubPlanJsonPreview = string.Empty;
    [ObservableProperty] private string epubResultJsonPreview = string.Empty;
    [ObservableProperty] private string epubLastPlanPath = string.Empty;
    [ObservableProperty] private string epubLastResultPath = string.Empty;

    public IReadOnlyList<int> EpubInputModeOptions { get; } = new[] { 0, 1 };

    public IRelayCommand BrowseEpubInputFilesCommand { get; private set; } = null!;
    public IRelayCommand RemoveEpubInputFileCommand { get; private set; } = null!;
    public IRelayCommand ClearEpubInputFilesCommand { get; private set; } = null!;
    public IRelayCommand BrowseEpubInputDirectoryCommand { get; private set; } = null!;
    public IRelayCommand BrowseEpubOutputDirectoryCommand { get; private set; } = null!;
    public IRelayCommand LoadLatestEpubArtifactsCommand { get; private set; } = null!;
    public IAsyncRelayCommand GenerateEpubPlanCommand { get; private set; } = null!;
    public IAsyncRelayCommand ExecuteEpubRunCommand { get; private set; } = null!;

    private StringBuilder _epubLogBuilder = new();
    public string EpubPipelineScriptPath => Path.Combine(_repoRoot, "Invoke-EpubUpscalePipeline.ps1");

    private void InitializeEpubPipelineFeatures()
    {
        EpubInputMode = 0;
        EpubRunStatusText = L("Status.Idle");
        EpubPreflightSummaryText = L("Status.NoPreflight");

        BrowseEpubInputFilesCommand = new RelayCommand(BrowseEpubInputFiles);
        RemoveEpubInputFileCommand = new RelayCommand(RemoveSelectedEpubInputFile, () => SelectedEpubInputFile is not null);
        ClearEpubInputFilesCommand = new RelayCommand(ClearEpubInputFiles);
        BrowseEpubInputDirectoryCommand = new RelayCommand(BrowseEpubInputDirectory);
        BrowseEpubOutputDirectoryCommand = new RelayCommand(BrowseEpubOutputDirectory);
        LoadLatestEpubArtifactsCommand = new RelayCommand(LoadLatestEpubArtifacts);
        GenerateEpubPlanCommand = new AsyncRelayCommand(GenerateEpubPlanAsync, () => !IsRunning);
        ExecuteEpubRunCommand = new AsyncRelayCommand(ExecuteEpubPipelineAsync, () => !IsRunning);
    }

    private void NotifyEpubPipelineCanExecuteChanged()
    {
        GenerateEpubPlanCommand.NotifyCanExecuteChanged();
        ExecuteEpubRunCommand.NotifyCanExecuteChanged();
        RemoveEpubInputFileCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedEpubInputFileChanged(string? value)
    {
        RemoveEpubInputFileCommand.NotifyCanExecuteChanged();
    }

    partial void OnEpubInputModeChanged(int value)
    {
        if (value == 0)
        {
            EpubInputDirectory = string.Empty;
        }
        else
        {
            EpubInputFiles.Clear();
            SelectedEpubInputFile = null;
        }
    }

    private void BrowseEpubInputFiles()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "EPUB (*.epub)|*.epub|All files (*.*)|*.*",
            InitialDirectory = GetInitialDirectory(EpubInputFiles.FirstOrDefault() ?? _repoRoot),
            CheckFileExists = true,
            Multiselect = true,
            Title = L("Msg.SelectEpubFiles")
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var existing = new HashSet<string>(EpubInputFiles, StringComparer.OrdinalIgnoreCase);
        foreach (var file in dialog.FileNames.Select(Path.GetFullPath).OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            if (existing.Add(file))
            {
                EpubInputFiles.Add(file);
            }
        }
        EpubInputMode = 0;
    }

    private void RemoveSelectedEpubInputFile()
    {
        if (SelectedEpubInputFile is null)
        {
            return;
        }
        EpubInputFiles.Remove(SelectedEpubInputFile);
        SelectedEpubInputFile = null;
    }

    private void ClearEpubInputFiles()
    {
        EpubInputFiles.Clear();
        SelectedEpubInputFile = null;
    }

    private void BrowseEpubInputDirectory()
    {
        var path = BrowseFolderPath(EpubInputDirectory, L("Msg.SelectEpubDirectory"));
        if (!string.IsNullOrWhiteSpace(path))
        {
            EpubInputDirectory = path;
            EpubInputMode = 1;
        }
    }

    private void BrowseEpubOutputDirectory()
    {
        var path = BrowseFolderPath(EpubOutputDirectory, L("Msg.SelectEpubOutputDirectory"));
        if (!string.IsNullOrWhiteSpace(path))
        {
            EpubOutputDirectory = path;
        }
    }

    private async Task GenerateEpubPlanAsync()
    {
        if (!ValidateEpubInputs(out var error))
        {
            ShowError(error);
            return;
        }
        await RunEpubPipelineProcessAsync(planOnly: true).ConfigureAwait(false);
    }

    private async Task ExecuteEpubPipelineAsync()
    {
        if (!ValidateEpubInputs(out var error))
        {
            ShowError(error);
            return;
        }

        var planRun = await RunEpubPipelineProcessAsync(planOnly: true).ConfigureAwait(false);
        if (planRun is null || planRun.WasCanceled)
        {
            return;
        }
        if (GetEpubPreflightCount("ERROR") > 0)
        {
            ShowError(L("Msg.EpubPreflightHasErrors"));
            return;
        }
        if (GetEpubPreflightCount("WARN") > 0)
        {
            var ok = AskYesNo(BuildEpubWarningMessage(), L("Msg.PreflightWarningsTitle"));
            if (!ok)
            {
                return;
            }
        }

        await RunEpubPipelineProcessAsync(planOnly: false).ConfigureAwait(false);
    }

    private bool ValidateEpubInputs(out string error)
    {
        if (!File.Exists(EpubPipelineScriptPath))
        {
            error = L("Msg.Validate.EpubScriptPath");
            return false;
        }
        if (string.IsNullOrWhiteSpace(DepsConfigPath) || !File.Exists(DepsConfigPath))
        {
            error = LF("Msg.Validate.DepsJsonFmt", "file not found");
            return false;
        }
        if (string.IsNullOrWhiteSpace(ConfigPath) || !File.Exists(ConfigPath))
        {
            error = LF("Msg.Validate.ConfigJsonFmt", "file not found");
            return false;
        }
        if (EpubInputMode == 0)
        {
            if (EpubInputFiles.Count < 1)
            {
                error = L("Msg.Validate.EpubInputFiles");
                return false;
            }
            if (EpubInputFiles.Any(x => !File.Exists(x)))
            {
                error = L("Msg.Validate.EpubInputFiles");
                return false;
            }
        }
        else
        {
            if (string.IsNullOrWhiteSpace(EpubInputDirectory) || !Directory.Exists(EpubInputDirectory))
            {
                error = L("Msg.Validate.EpubInputDirectory");
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(EpubOutputDirectory))
        {
            try
            {
                var full = Path.GetFullPath(EpubOutputDirectory);
                if (File.Exists(full))
                {
                    error = L("Msg.Validate.EpubOutputDirectory");
                    return false;
                }
            }
            catch
            {
                error = L("Msg.Validate.EpubOutputDirectory");
                return false;
            }
        }
        if (UpscaleFactor < 1 || UpscaleFactor > 4)
        {
            error = L("Run.Tip.UpscaleFactor");
            return false;
        }
        if (LossyQuality < 1 || LossyQuality > 100)
        {
            error = L("Run.Tip.LossyQuality");
            return false;
        }
        if (CfgGrayscaleDetectionThreshold < 0 || CfgGrayscaleDetectionThreshold > 24)
        {
            error = L("Cfg.Tip.GrayscaleDetectionThreshold");
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
        error = string.Empty;
        return true;
    }

    private List<string> BuildEpubPipelineArguments(bool planOnly)
    {
        var args = new List<string>
        {
            "-GuiMode",
            "-AutoConfirm",
            "-UpscaleFactor", UpscaleFactor.ToString(CultureInfo.InvariantCulture),
            "-LossyQuality", LossyQuality.ToString(CultureInfo.InvariantCulture),
            "-GrayscaleDetectionThreshold", CfgGrayscaleDetectionThreshold.ToString(CultureInfo.InvariantCulture),
            "-LogLevel", LogLevel
        };
        var suffix = (EpubOutputSuffix ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(suffix))
        {
            args.Add("-NoOutputSuffix");
        }
        else
        {
            args.Add("-OutputSuffix");
            args.Add(suffix);
        }
        if (!string.IsNullOrWhiteSpace(EpubOutputDirectory))
        {
            args.Add("-OutputDirectory");
            args.Add(EpubOutputDirectory);
        }

        if (EpubInputMode == 0)
        {
            var sortedFiles = EpubInputFiles
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            args.Add("-InputEpubPath");
            // powershell.exe -File cannot reliably bind a native multi-value array argument;
            // pass one delimiter-joined payload and let the script split it.
            args.Add(string.Join("|", sortedFiles));
        }
        else
        {
            args.Add("-InputEpubDirectory");
            args.Add(EpubInputDirectory);
        }

        if (!string.IsNullOrWhiteSpace(DepsConfigPath)) { args.Add("-DepsConfigPath"); args.Add(DepsConfigPath); }
        if (!string.IsNullOrWhiteSpace(ConfigPath)) { args.Add("-ConfigPath"); args.Add(ConfigPath); }
        if (EpubOverwriteOutputEpub) { args.Add("-OverwriteOutputEpub"); }
        if (DryRun) { args.Add("-DryRun"); }
        if (FailOnPreflightWarnings) { args.Add("-FailOnPreflightWarnings"); }
        if (planOnly) { args.Add("-PlanOnly"); }

        return args;
    }

    private async Task<ProcessRunResult?> RunEpubPipelineProcessAsync(bool planOnly)
    {
        if (IsRunning)
        {
            return null;
        }

        RunOnUi(() =>
        {
            IsRunning = true;
            EpubRunStatusText = planOnly ? L("Status.RunningPlan") : L("Status.RunningPipeline");
            if (!planOnly)
            {
                EpubFilePercent = 0;
                EpubFileStatusText = "-";
            }
        });

        _runningCts = new CancellationTokenSource();
        var args = BuildEpubPipelineArguments(planOnly);
        AppendEpubLog($"Run: {string.Join(" ", args)}");

        var stderrTail = new List<string>();
        ProcessRunResult? result = null;
        try
        {
            result = await _processRunner.RunPowerShellFileAsync(
                EpubPipelineScriptPath,
                args,
                line => RunOnUi(() => AppendEpubLog(line)),
                line => RunOnUi(() =>
                {
                    AppendEpubLog("[stderr] " + line);
                    if (stderrTail.Count >= 20)
                    {
                        stderrTail.RemoveAt(0);
                    }
                    stderrTail.Add(line);
                }),
                ev => RunOnUi(() => HandleEpubPipelineEvent(ev)),
                _runningCts.Token).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            RunOnUi(() => ShowError(LF("Msg.LaunchFailedFmt", ex.Message)));
        }

        RunOnUi(() =>
        {
            if (result is not null)
            {
                EpubRunStatusText = result.WasCanceled ? L("Status.Canceled") : LF("Status.ExitFmt", result.ExitCode);
                AppendEpubLog(LF("Msg.ProcessExitFmt", result.ExitCode));
                LoadLatestEpubArtifacts();
                if (!result.WasCanceled && result.ExitCode != 0)
                {
                    ShowError(BuildEpubFailureMessage(result.ExitCode, stderrTail));
                }
            }
            IsRunning = false;
        });
        _runningCts?.Dispose();
        _runningCts = null;
        return result;
    }

    private void HandleEpubPipelineEvent(PipelineEventEnvelope ev)
    {
        switch (ev.Type)
        {
            case "plan_ready":
                EpubLastPlanPath = GetString(ev.Data, "plan_path");
                break;
            case "preflight_summary":
                ParseEpubPreflight(ev.Data);
                break;
            case "epub_file_progress":
                ParseEpubFileProgress(ev.Data);
                break;
            case "run_result":
                EpubLastResultPath = GetString(ev.Data, "result_path");
                break;
            default:
                AppendEpubLog(LF("Msg.EventFmt", ev.Type, ev.RawJson));
                break;
        }
    }

    private void ParseEpubPreflight(JsonElement data)
    {
        var errors = GetInt(data, "errors");
        var warnings = GetInt(data, "warnings");
        var infos = GetInt(data, "infos");
        EpubPreflightSummaryText = LF("Run.PreflightSummaryFmt", errors, warnings, infos);
        EpubPreflightIssues.Clear();

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

                EpubPreflightIssues.Add(new PipelineIssue
                {
                    Severity = NormalizeSeverity(severity),
                    Code = code,
                    Message = message
                });
            }
        }
    }

    private void ParseEpubFileProgress(JsonElement data)
    {
        var done = GetInt(data, "done_files");
        var total = Math.Max(1, GetInt(data, "total_files"));
        var fileDoneUnits = Math.Max(0, GetInt(data, "file_done_units"));
        var fileTotalUnits = Math.Max(0, GetInt(data, "file_total_units"));
        var percent = GetDouble(data, "percent");
        if (percent <= 0)
        {
            percent = (done * 100.0) / total;
        }

        EpubFilePercent = Math.Clamp(percent, 0, 100);
        var status = GetString(data, "status");
        var current = GetString(data, "current_file");
        var imagePart = fileTotalUnits > 0 ? $" img {fileDoneUnits}/{fileTotalUnits}" : string.Empty;
        EpubFileStatusText = string.IsNullOrWhiteSpace(current)
            ? $"{done}/{total}{imagePart} [{status}]"
            : $"{done}/{total}{imagePart} [{status}] {current}";
    }

    private int GetEpubPreflightCount(string severity)
    {
        var expected = NormalizeSeverity(severity);
        return EpubPreflightIssues.Count(x => string.Equals(NormalizeSeverity(x.Severity), expected, StringComparison.OrdinalIgnoreCase));
    }

    private string BuildEpubWarningMessage()
    {
        var lines = EpubPreflightIssues
            .Where(x => string.Equals(NormalizeSeverity(x.Severity), "WARN", StringComparison.OrdinalIgnoreCase))
            .Select(x => $"- [{x.Code}] {x.Message}")
            .ToList();

        if (lines.Count == 0) return L("Msg.PreflightWarningsContinue");
        return L("Msg.PreflightWarningsHeader") + "\n\n" + string.Join("\n", lines) + "\n\n" + L("Msg.PreflightWarningsContinue");
    }

    private string BuildEpubFailureMessage(int exitCode, IReadOnlyList<string> stderrTail)
    {
        var lines = new List<string>
        {
            LF("Msg.ProcessExitFmt", exitCode)
        };

        if (!string.IsNullOrWhiteSpace(EpubLastResultPath))
        {
            lines.Add("latest_epub_run_result.json: " + EpubLastResultPath);
        }

        var latestLogPath = GetLatestEpubRunLogPath();
        if (!string.IsNullOrWhiteSpace(latestLogPath))
        {
            lines.Add("log: " + latestLogPath);
        }

        if (stderrTail.Count > 0)
        {
            lines.Add(string.Empty);
            lines.Add("stderr:");
            foreach (var item in stderrTail.Skip(Math.Max(0, stderrTail.Count - 8)))
            {
                lines.Add("  " + item);
            }
        }

        return string.Join(Environment.NewLine, lines);
    }

    private string GetLatestEpubRunLogPath()
    {
        try
        {
            var logDir = Path.Combine(_repoRoot, "logs");
            if (!Directory.Exists(logDir))
            {
                return string.Empty;
            }

            var latest = new DirectoryInfo(logDir)
                .GetFiles("epub_upscale_run_*.log", SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .FirstOrDefault();

            return latest?.FullName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void AppendEpubLog(string line)
    {
        _epubLogBuilder.AppendLine(line);
        EpubLiveLogText = _epubLogBuilder.ToString();
        CaptureEpubRuntimeWarning(line);
    }

    private void CaptureEpubRuntimeWarning(string line)
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

        var exists = EpubPreflightIssues.Any(x =>
            string.Equals(NormalizeSeverity(x.Severity), "WARN", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.Code, code, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.Message, message, StringComparison.Ordinal));
        if (!exists)
        {
            EpubPreflightIssues.Add(new PipelineIssue
            {
                Severity = "WARN",
                Code = code,
                Message = message
            });
        }
    }

    private void LoadLatestEpubArtifacts()
    {
        var planPath = Path.Combine(_repoRoot, "logs", "latest_epub_run_plan.json");
        if (File.Exists(planPath))
        {
            EpubLastPlanPath = planPath;
            var text = File.ReadAllText(planPath, Encoding.UTF8);
            EpubPlanJsonPreview = text;
            TryPopulateEpubPreflightFromPlan(text);
        }

        var resultPath = Path.Combine(_repoRoot, "logs", "latest_epub_run_result.json");
        if (File.Exists(resultPath))
        {
            EpubLastResultPath = resultPath;
            EpubResultJsonPreview = File.ReadAllText(resultPath, Encoding.UTF8);
        }
    }

    private void TryPopulateEpubPreflightFromPlan(string planJson)
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

            EpubPreflightIssues.Clear();
            if (TryGetProperty(preflight, "Issues", out var issues) && issues.ValueKind == JsonValueKind.Array)
            {
                foreach (var it in issues.EnumerateArray())
                {
                    var severity = GetString(it, "Severity");
                    var code = GetString(it, "Code");
                    var message = GetString(it, "Message");
                    EpubPreflightIssues.Add(new PipelineIssue
                    {
                        Severity = NormalizeSeverity(severity),
                        Code = code,
                        Message = message
                    });
                }
            }

            EpubPreflightSummaryText = LF("Run.PreflightSummaryFmt", errors, warnings, infos);
        }
        catch
        {
            // Keep previous UI state when latest plan JSON can't be parsed.
        }
    }
}
