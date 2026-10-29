namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Conservative JSON-size admission, including escaped text and collection structure.</summary>
internal sealed class RangeReadBudget
{
    internal const long MaximumResponseBytes = 4 * 1024 * 1024;
    // MCP carries structured content and JSON text, whose own escaping can add another copy.
    private const int TransportCopies = 3;
    internal long EstimatedBytes { get; private set; } = 4096 * TransportCopies;

    internal void ReserveGrid(int rows, int columns, bool includeFormulas) =>
        Add((long)rows * (32L + 16L * columns) * (includeFormulas ? 2 : 1) * TransportCopies);

    internal void AddScalar(object? value) => Add(EstimateScalar(value) * TransportCopies);

    internal void AddError(string formula) => Add((4096 + EstimateScalar(formula)) * TransportCopies);

    internal static long EstimateScalar(object? value) => value switch
    {
        string text => 2L + 6L * text.Length,
        null or DBNull => 4,
        bool => 5,
        _ => 64
    };

    private void Add(long bytes)
    {
        if (bytes > MaximumResponseBytes - EstimatedBytes)
            throw new ArgumentOutOfRangeException(nameof(bytes),
                "The requested data exceeds the conservative 4 MiB response byte budget. Read smaller explicit ranges or reduce pageRows/pageColumns. No partial values were returned.");
        EstimatedBytes += bytes;
    }
}
