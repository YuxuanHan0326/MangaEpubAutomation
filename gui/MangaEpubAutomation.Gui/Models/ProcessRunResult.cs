namespace MangaEpubAutomation.Gui.Models;

public sealed record ProcessRunResult(
    int ExitCode,
    bool WasCanceled
);
