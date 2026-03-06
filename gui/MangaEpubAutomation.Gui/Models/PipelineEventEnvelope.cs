using System.Text.Json;

namespace MangaEpubAutomation.Gui.Models;

public sealed record PipelineEventEnvelope(
    string Type,
    string? TimestampUtc,
    JsonElement Data,
    string RawJson
);
