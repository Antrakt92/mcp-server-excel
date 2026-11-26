using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.Security;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service;

internal enum ExcelCrashEventStatus
{
    [JsonStringEnumMemberName("matched")]
    Matched,
    [JsonStringEnumMemberName("notFound")]
    NotFound,
    [JsonStringEnumMemberName("ambiguous")]
    Ambiguous,
    [JsonStringEnumMemberName("unavailable")]
    Unavailable,
    [JsonStringEnumMemberName("notApplicable")]
    NotApplicable
}

internal sealed record ExcelCrashEventResult(ExcelCrashEventStatus Status, ExcelCrashEventDetails? Event = null);
internal sealed record ExcelCrashEventDetails(DateTimeOffset TimeUtc, string? ApplicationVersion,
    string? ModuleName, string? ModuleOffset, string? ExceptionCode);

internal sealed record ExcelCrashEventCandidate(DateTimeOffset TimeUtc, string? ApplicationName, string? ApplicationVersion,
    string? ModuleName, string? ModuleOffset, string? ExceptionCode, object? ProcessId, object? ProcessCreationTime);

/// <summary>Looks up a Windows Application Error event for an already observed session process exit.</summary>
internal static class ExcelCrashEventReader
{
    internal const int MaximumEvents = 64;
    private static readonly TimeSpan ReadBudget = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan BeforeExit = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan AfterExit = TimeSpan.FromMinutes(5);
    private static readonly string[] SelectedFields =
    [
        "Event/EventData/Data[@Name='AppName']",
        "Event/EventData/Data[@Name='AppVersion']",
        "Event/EventData/Data[@Name='ModuleName']",
        "Event/EventData/Data[@Name='FaultingOffset']",
        "Event/EventData/Data[@Name='ExceptionCode']",
        "Event/EventData/Data[@Name='ProcessId']",
        "Event/EventData/Data[@Name='ProcessCreationTime']"
    ];

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An optional local event lookup reports unavailable if the log cannot be accessed or searched within its budget. The already observed native exit is independent.")]
    internal static ExcelCrashEventResult Read(ExcelNativeProcessObservation observation, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (QueryEligibility(observation) is { } reason)
            return new(reason);

        try
        {
            return Match(observation, ReadCandidates(observation.ExitedAtUtc!.Value, cancellationToken), cancellationToken);
        }
        catch (Exception error) when (error is EventLogException or UnauthorizedAccessException or SecurityException or TimeoutException)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return new(ExcelCrashEventStatus.Unavailable);
        }
    }

    internal static ExcelCrashEventResult Match(ExcelNativeProcessObservation observation,
        IEnumerable<ExcelCrashEventCandidate> candidates, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (QueryEligibility(observation) is { } reason)
            return new(reason);

        var (earliest, latest) = EventWindow(observation.ExitedAtUtc!.Value);
        ExcelCrashEventDetails? match = null;
        var scanned = 0;
        foreach (var candidate in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++scanned > MaximumEvents)
                return new(ExcelCrashEventStatus.Unavailable);
            if (!string.Equals(candidate.ApplicationName, "EXCEL.EXE", StringComparison.OrdinalIgnoreCase)
                || candidate.TimeUtc < earliest || candidate.TimeUtc > latest
                || ReadIdentityNumber(candidate.ProcessId) != (ulong)observation.Identity.ProcessId
                || ReadIdentityNumber(candidate.ProcessCreationTime) != (ulong)observation.Identity.StartedAtUtcFileTime)
                continue;

            var detail = new ExcelCrashEventDetails(candidate.TimeUtc.ToUniversalTime(),
                ReadApplicationVersion(candidate.ApplicationVersion), ReadModuleName(candidate.ModuleName),
                ReadHex(candidate.ModuleOffset, 16), ReadHex(candidate.ExceptionCode, 8));
            if (match is not null && match != detail)
                return new(ExcelCrashEventStatus.Ambiguous);
            match = detail;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return match is null ? new(ExcelCrashEventStatus.NotFound) : new(ExcelCrashEventStatus.Matched, match);
    }

    private static ExcelCrashEventStatus? QueryEligibility(ExcelNativeProcessObservation observation)
    {
        if (observation.State != ExcelNativeProcessState.Exited || observation.ExitCode is null or 0)
            return ExcelCrashEventStatus.NotApplicable;
        if (observation.ExitedAtUtc is not { } exitTime
            || observation.Identity is { ProcessId: <= 0 } or { StartedAtUtcFileTime: <= 0 }
            || exitTime - DateTimeOffset.MinValue < BeforeExit || DateTimeOffset.MaxValue - exitTime < AfterExit)
            return ExcelCrashEventStatus.Unavailable;
        return null;
    }

    // Event 1000 can be published after the native process has exited. Search
    // relative to the native exit, and let an explicit caller retry a miss.
    private static (DateTimeOffset Earliest, DateTimeOffset Latest) EventWindow(DateTimeOffset exitTime) =>
        (exitTime - BeforeExit, exitTime + AfterExit);

    private static IEnumerable<ExcelCrashEventCandidate> ReadCandidates(DateTimeOffset exitedAtUtc,
        CancellationToken cancellationToken)
    {
        var (earliest, latest) = EventWindow(exitedAtUtc);
        var query = FormattableString.Invariant(
            $"*[System[Provider[@Name='Application Error'] and EventID=1000 and TimeCreated[@SystemTime >= '{earliest.UtcDateTime:O}' and @SystemTime <= '{latest.UtcDateTime:O}']] and EventData[Data[@Name='AppName']='EXCEL.EXE' or Data[@Name='AppName']='Excel.exe' or Data[@Name='AppName']='excel.exe']]");
        var started = Stopwatch.GetTimestamp();
        using var reader = new EventLogReader(new EventLogQuery("Application", PathType.LogName, query));
        using var selector = new EventLogPropertySelector(SelectedFields);
        var scanned = 0;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remaining = ReadBudget - Stopwatch.GetElapsedTime(started);
            if (remaining <= TimeSpan.Zero)
                throw new TimeoutException("The local event lookup exceeded its query budget.");

            using var record = reader.ReadEvent(remaining);
            cancellationToken.ThrowIfCancellationRequested();
            if (Stopwatch.GetElapsedTime(started) > ReadBudget)
                throw new TimeoutException("The local event lookup exceeded its query budget.");
            if (record is null)
                yield break;
            if (++scanned > MaximumEvents)
                throw new TimeoutException("The local event lookup exceeded its event limit.");
            if (record is not EventLogRecord entry || entry.TimeCreated is not DateTime time)
                continue;

            var values = entry.GetPropertyValues(selector);
            if (values.Count != SelectedFields.Length)
                continue;
            yield return new(time.ToUniversalTime(), values[0] as string, values[1] as string, values[2] as string,
                values[3] as string, values[4] as string, values[5], values[6]);
        }
    }

    private static ulong? ReadIdentityNumber(object? value) => value switch
    {
        uint number => number,
        ulong number => number,
        int number when number > 0 => (ulong)number,
        long number when number > 0 => (ulong)number,
        string text when text.StartsWith("0x", StringComparison.OrdinalIgnoreCase) =>
            ulong.TryParse(text.AsSpan(2), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var hex) ? hex : null,
        string text => ulong.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var dec) ? dec : null,
        _ => null
    };

    private static string? ReadApplicationVersion(string? value) =>
        value is { Length: > 0 and <= 48 } && Version.TryParse(value, out var version) ? version.ToString() : null;

    private static string? ReadModuleName(string? value) =>
        value is { Length: > 0 and <= 128 } && value is not "." and not ".."
            && value.All(character => char.IsAsciiLetterOrDigit(character) || character is '.' or '-' or '_' or ' ') ? value : null;

    private static string? ReadHex(string? value, int digits)
    {
        if (value is null)
            return null;
        var hex = value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? value.AsSpan(2) : value.AsSpan();
        if (hex.Length == 0 || hex.Length > digits
            || !ulong.TryParse(hex, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var parsed))
            return null;
        return "0x" + parsed.ToString($"X{digits}", CultureInfo.InvariantCulture);
    }
}
