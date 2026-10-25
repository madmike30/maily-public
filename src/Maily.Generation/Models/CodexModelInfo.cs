namespace Maily.Generation.Models;

public sealed record CodexModelInfo(string Id, string DisplayName, string Description,
    IReadOnlyList<string> ReasoningEfforts, bool IsDefault);

public sealed record CodexModelCatalogSnapshot(string Status, string Source,
    IReadOnlyList<CodexModelInfo> Models, string? RecommendedModel, string SelectionReason,
    DateTimeOffset FetchedAtUtc, string? Warning);
