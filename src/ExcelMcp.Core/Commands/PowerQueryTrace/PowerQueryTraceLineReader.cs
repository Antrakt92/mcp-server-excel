using System.Text;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

internal sealed class PowerQueryTraceReadBudget(int characters, int lines)
{
    public int RemainingCharacters { get; set; } = characters;
    public int RemainingLines { get; set; } = lines;
}

internal static class PowerQueryTraceLineReader
{
    internal const int MaxLineCharacters = 256 * 1024;

    /// <summary>Reads complete physical lines without allocating an arbitrarily large string.</summary>
    internal static IEnumerable<string?> Read(
        string path,
        PowerQueryTraceReadBudget budget,
        int fileCharacters,
        int fileLines,
        Action reportScanLimit,
        Action reportLineLimit)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = new StreamReader(stream);
        var buffer = new char[4096];
        var line = new StringBuilder();
        bool oversized = false;
        bool skipLineFeed = false;

        while (budget.RemainingCharacters > 0 && fileCharacters > 0)
        {
            int read = reader.ReadBlock(buffer, 0, Math.Min(buffer.Length, Math.Min(budget.RemainingCharacters, fileCharacters)));
            if (read == 0)
                break;
            budget.RemainingCharacters -= read;
            fileCharacters -= read;

            int offset = 0;
            while (offset < read)
            {
                if (skipLineFeed && buffer[offset] == '\n')
                    offset++;
                skipLineFeed = false;
                if (offset == read)
                    break;
                if (budget.RemainingLines <= 0 || fileLines <= 0)
                {
                    reportScanLimit();
                    yield break;
                }

                int newline = buffer.AsSpan(offset, read - offset).IndexOfAny('\r', '\n');
                int length = newline < 0 ? read - offset : newline;
                if (!oversized && line.Length + length > MaxLineCharacters)
                {
                    oversized = true;
                    line.Clear();
                    reportLineLimit();
                }
                if (!oversized)
                    line.Append(buffer, offset, length);
                offset += length;
                if (newline < 0)
                    continue;

                skipLineFeed = buffer[offset++] == '\r';
                budget.RemainingLines--;
                fileLines--;
                yield return oversized ? null : line.ToString();
                line.Clear();
                oversized = false;
            }
        }

        if (reader.Peek() != -1)
        {
            reportScanLimit();
            yield break;
        }
        if (line.Length > 0 || oversized)
        {
            budget.RemainingLines--;
            yield return oversized ? null : line.ToString();
        }
    }
}
