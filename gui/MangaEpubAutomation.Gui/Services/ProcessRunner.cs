using System.Diagnostics;
using System.Text;
using System.Text.Json;
using MangaEpubAutomation.Gui.Models;

namespace MangaEpubAutomation.Gui.Services;

public sealed class ProcessRunner
{
    private const string PipelineEventPrefix = "PIPELINE_EVENT:";

    public async Task<ProcessRunResult> RunPowerShellFileAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        Action<string>? onStdout,
        Action<string>? onStderr,
        Action<PipelineEventEnvelope>? onEvent,
        CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        startInfo.ArgumentList.Add("-ExecutionPolicy");
        startInfo.ArgumentList.Add("Bypass");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(scriptPath);

        foreach (var arg in arguments)
        {
            startInfo.ArgumentList.Add(arg);
        }

        using var process = new Process
        {
            StartInfo = startInfo,
            EnableRaisingEvents = true
        };

        var exitTcs = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        var wasCanceled = false;

        process.Exited += (_, _) =>
        {
            exitTcs.TrySetResult(process.ExitCode);
        };

        process.OutputDataReceived += (_, e) =>
        {
            var line = e.Data;
            if (string.IsNullOrEmpty(line))
            {
                return;
            }

            if (TryParsePipelineEvent(line, out var envelope))
            {
                try
                {
                    onEvent?.Invoke(envelope);
                }
                catch (Exception ex)
                {
                    onStderr?.Invoke("[gui-event-callback-error] " + ex.Message);
                }
                return;
            }

            try
            {
                onStdout?.Invoke(line);
            }
            catch (Exception ex)
            {
                onStderr?.Invoke("[gui-stdout-callback-error] " + ex.Message);
            }
        };

        process.ErrorDataReceived += (_, e) =>
        {
            var line = e.Data;
            if (string.IsNullOrEmpty(line))
            {
                return;
            }

            try
            {
                onStderr?.Invoke(line);
            }
            catch
            {
                // Ignore stderr callback exceptions.
            }
        };

        if (!process.Start())
        {
            throw new InvalidOperationException("Failed to start pipeline process.");
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var registration = cancellationToken.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                {
                    wasCanceled = true;
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // Ignore kill races.
            }
        });

        var waitForExitTask = process.WaitForExitAsync();
        await Task.WhenAny(exitTcs.Task, waitForExitTask).ConfigureAwait(false);
        var exitCode = process.HasExited ? process.ExitCode : await exitTcs.Task.ConfigureAwait(false);

        try
        {
            process.WaitForExit();
        }
        catch
        {
            // Best effort.
        }

        return new ProcessRunResult(exitCode, wasCanceled);
    }

    private static bool TryParsePipelineEvent(string line, out PipelineEventEnvelope envelope)
    {
        envelope = null!;
        if (!line.StartsWith(PipelineEventPrefix, StringComparison.Ordinal))
        {
            return false;
        }

        var payload = line[PipelineEventPrefix.Length..].Trim();
        if (payload.Length == 0)
        {
            return false;
        }

        try
        {
            using var doc = JsonDocument.Parse(payload);
            var root = doc.RootElement;

            var type = root.TryGetProperty("type", out var typeElement)
                ? typeElement.GetString() ?? string.Empty
                : string.Empty;

            var ts = root.TryGetProperty("ts_utc", out var tsElement)
                ? tsElement.GetString()
                : null;

            JsonElement data = default;
            if (root.TryGetProperty("data", out var dataElement))
            {
                data = dataElement.Clone();
            }

            envelope = new PipelineEventEnvelope(type, ts, data, payload);
            return true;
        }
        catch
        {
            // If malformed, pass through as normal output line.
            return false;
        }
    }
}
