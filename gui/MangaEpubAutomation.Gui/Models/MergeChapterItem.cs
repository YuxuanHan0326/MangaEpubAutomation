namespace MangaEpubAutomation.Gui.Models;

public sealed class MergeChapterItem
{
    public int Index { get; set; }

    public string Chapter { get; set; } = string.Empty;

    public string Order { get; set; } = string.Empty;

    public string EpubPath { get; set; } = string.Empty;
}
