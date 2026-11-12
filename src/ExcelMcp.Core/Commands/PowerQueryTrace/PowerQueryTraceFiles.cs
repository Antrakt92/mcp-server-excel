using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

internal static class PowerQueryTraceFiles
{
    public static int Count(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return 0;
        }

        return EnumerateTraceFiles(folder).Count();
    }

    public static IEnumerable<FileInfo> EnumerateTraceFiles(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder)
            .Select(path => new FileInfo(path))
            .Where(file => PowerQueryTraceFileName.TryParse(file.Name, out _, out _, out _));
    }
}

/// <summary>
/// Reads local Power Query trace files and produces file lists and summaries.
/// </summary>
internal sealed class PowerQueryTraceReader
{
    private const int ContainerHostPeekLineCount = 50;
    private const int ContainerHostPeekCharacters = 256 * 1024;
    private const int HeaderScanCharacters = 8 * 1024 * 1024;
    private const int MaxScannedFiles = 10_000;
    private const int MaxListedFiles = 512;
    private const int MaxSummarizedFiles = 256;
    private const int SummaryScanCharacters = 32 * 1024 * 1024;
    private const int SummaryScanLines = 100_000;
    private const string HeaderLimit = "Container host detection reached its 50-line, 256 Ki-character per-file, or 8 Mi-character total scan limit. Some containers could not be attributed.";
    private const string LineLimit = "At least one trace line exceeded the 256 Ki-character line limit and was skipped.";
    private const string SummaryLimit = "The summary reached its 32 Mi-character or 100,000-line scan limit.";
    private readonly string _traceFolder;
    private readonly Func<int, bool> _isProcessRunning;

    public PowerQueryTraceReader(string traceFolder, Func<int, bool>? isProcessRunning = null)
    {
        _traceFolder = traceFolder ?? throw new ArgumentNullException(nameof(traceFolder));
        _isProcessRunning = isProcessRunning ?? IsProcessRunning;
    }

    public PowerQueryTraceListResult List(
        int? processId,
        DateTime? sinceUtc,
        DateTime? processStartedUtc = null)
    {
        var files = new List<PowerQueryTraceFileInfo>();
        var warnings = new HashSet<string>(StringComparer.Ordinal);
        var headers = new PowerQueryTraceReadBudget(HeaderScanCharacters, int.MaxValue);
        int matchingFileCount = 0;
        long totalSizeBytes = 0;
        foreach (var (file, kind, fileProcessId, startedUtc) in ScanTraceFiles(warnings))
        {
            if (processStartedUtc is DateTime processStart &&
                (startedUtc is not DateTime fileStart || fileStart < processStart))
            {
                continue;
            }

            int? hostProcessId = null;
            if (processId is int requestedProcessId)
            {
                if (kind == PowerQueryTraceFileName.ExcelKind)
                {
                    if (fileProcessId != requestedProcessId)
                    {
                        continue;
                    }
                }
                else
                {
                    hostProcessId = FindHostProcessId(file.FullName, requestedProcessId, headers, warnings);
                    if (hostProcessId != requestedProcessId)
                    {
                        continue;
                    }
                }
            }
            else if (kind == PowerQueryTraceFileName.ContainerKind && sinceUtc.HasValue)
            {
                hostProcessId = FindHostProcessId(file.FullName, expectedProcessId: null, headers, warnings);
            }

            DateTime lastWriteUtc = file.LastWriteTimeUtc;
            if (sinceUtc is DateTime since &&
                lastWriteUtc < since &&
                (startedUtc is not DateTime started || started < since) &&
                !_isProcessRunning(fileProcessId) &&
                (hostProcessId is not int hostPid || !_isProcessRunning(hostPid)))
            {
                continue;
            }

            long sizeBytes = file.Length;
            matchingFileCount++;
            totalSizeBytes += sizeBytes;
            if (files.Count >= MaxListedFiles)
                continue;
            files.Add(new PowerQueryTraceFileInfo
            {
                FileName = file.Name,
                Kind = kind,
                ProcessId = fileProcessId,
                HostProcessId = hostProcessId,
                StartedUtc = startedUtc,
                LastWriteUtc = lastWriteUtc,
                SizeBytes = sizeBytes
            });
        }

        if (matchingFileCount > MaxListedFiles)
            warnings.Add("The file list shows the 512 newest matching files from the scan.");
        return new PowerQueryTraceListResult
        {
            Success = true,
            TraceFolder = _traceFolder,
            ProcessId = processId,
            Files = files,
            TotalSizeBytes = totalSizeBytes,
            MatchingFileCount = matchingFileCount,
            Truncated = warnings.Count > 0,
            ScanWarnings = warnings.Count == 0 ? null : warnings.ToList()
        };
    }

    public PowerQueryTraceSummaryResult Summarize(PowerQueryTraceFilter filter)
    {
        var analyzer = new PowerQueryTraceAnalyzer(filter);
        var warnings = new HashSet<string>(StringComparer.Ordinal);
        var headers = new PowerQueryTraceReadBudget(HeaderScanCharacters, int.MaxValue);
        var budget = new PowerQueryTraceReadBudget(SummaryScanCharacters, SummaryScanLines);
        int filesRead = 0;
        foreach (var file in SelectFilesForSummary(filter, headers, warnings))
        {
            if (filesRead >= MaxSummarizedFiles)
            {
                warnings.Add("The summary reached its 256-file limit.");
                break;
            }
            if (budget.RemainingCharacters <= 0 || budget.RemainingLines <= 0)
            {
                warnings.Add(SummaryLimit);
                break;
            }
            filesRead++;
            analyzer.AddFileRead();
            foreach (var line in PowerQueryTraceLineReader.Read(file.FullName, budget, int.MaxValue, int.MaxValue,
                () => warnings.Add(SummaryLimit), () => warnings.Add(LineLimit)))
            {
                if (PowerQueryTraceLineParser.TryParse(line, out var entry) && entry != null)
                {
                    analyzer.Add(entry);
                }
                else
                {
                    analyzer.AddUnreadableLine();
                }
            }
        }

        var summary = analyzer.Build(truncated: warnings.Count > 0);
        if (summary.Truncated)
            summary.Findings.InsertRange(0, warnings);
        return summary;
    }

    private IEnumerable<FileInfo> SelectFilesForSummary(PowerQueryTraceFilter filter,
        PowerQueryTraceReadBudget headers, HashSet<string> warnings)
    {
        foreach (var (file, kind, fileProcessId, startedUtc) in ScanTraceFiles(warnings))
        {
            if (filter.UntilUtc is DateTime until && startedUtc is DateTime start && start > until)
            {
                continue;
            }

            if (filter.ProcessStartedUtc is DateTime processStart &&
                (startedUtc is not DateTime fileStart || fileStart < processStart))
            {
                continue;
            }

            int? hostProcessId = null;
            if (filter.ProcessId is int processId)
            {
                if (kind == PowerQueryTraceFileName.ExcelKind && fileProcessId != processId)
                {
                    continue;
                }

                if (kind == PowerQueryTraceFileName.ContainerKind)
                {
                    hostProcessId = FindHostProcessId(file.FullName, processId, headers, warnings);
                    if (hostProcessId != processId)
                    {
                        continue;
                    }
                }
            }

            if (filter.SinceUtc is DateTime since &&
                file.LastWriteTimeUtc < since &&
                (startedUtc is not DateTime started || started < since))
            {
                hostProcessId ??= kind == PowerQueryTraceFileName.ContainerKind
                    ? FindHostProcessId(file.FullName, expectedProcessId: null, headers, warnings)
                    : null;
                if (!_isProcessRunning(fileProcessId) &&
                    (hostProcessId is not int hostPid || !_isProcessRunning(hostPid)))
                {
                    continue;
                }
            }

            yield return file;
        }
    }

    private IEnumerable<(FileInfo File, string Kind, int ProcessId, DateTime? StartedUtc)> ScanTraceFiles(HashSet<string> warnings)
    {
        if (!Directory.Exists(_traceFolder))
            return [];

        var candidates = new List<(FileInfo File, string Kind, int ProcessId, DateTime? StartedUtc)>();
        int scanned = 0;
        foreach (var file in new DirectoryInfo(_traceFolder).EnumerateFiles())
        {
            if (++scanned > MaxScannedFiles)
            {
                warnings.Add("The trace directory scan reached its 10,000-file limit. Older or newer files may be missing.");
                break;
            }
            if (PowerQueryTraceFileName.TryParse(file.Name, out var kind, out var processId, out var startedUtc))
                candidates.Add((file, kind, processId, startedUtc));
        }
        return candidates.OrderByDescending(item => item.File.LastWriteTimeUtc)
            .ThenBy(item => item.File.Name, StringComparer.Ordinal);
    }

    private static int? FindHostProcessId(string filePath, int? expectedProcessId,
        PowerQueryTraceReadBudget headers, HashSet<string> warnings)
    {
        if (headers.RemainingCharacters <= 0)
        {
            warnings.Add(HeaderLimit);
            return null;
        }
        foreach (var line in PowerQueryTraceLineReader.Read(filePath, headers, ContainerHostPeekCharacters,
            ContainerHostPeekLineCount, () => warnings.Add(HeaderLimit), () => warnings.Add(LineLimit)))
        {
            if (PowerQueryTraceLineParser.TryParse(line, out var entry) &&
                entry?.HostProcessId is int hostProcessId &&
                (expectedProcessId is null || hostProcessId == expectedProcessId))
            {
                return hostProcessId;
            }
        }

        return null;
    }

    private static bool IsProcessRunning(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }
}
