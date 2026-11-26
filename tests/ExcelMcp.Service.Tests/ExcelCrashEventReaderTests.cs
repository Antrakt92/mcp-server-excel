using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelCrashEventReaderTests
{
    private static readonly DateTimeOffset ExitTime = new(2026, 1, 1, 12, 15, 0, TimeSpan.Zero);
    private static readonly ExcelProcessIdentity Identity = new(1234, ExitTime.AddMinutes(-30).ToFileTime());
    private static readonly ExcelNativeProcessObservation Exit = new(Identity, ExcelNativeProcessState.Exited,
        0xC0000005, null, ExitTime);

    [Fact]
    public void Match_SkipsReusedPidAndOtherApplicationsAndReturnsOnlyAllowedFields()
    {
        var otherStart = MakeCandidate(Identity.StartedAtUtcFileTime - 1);
        var otherApp = MakeCandidate() with { ApplicationName = "notexcel.exe" };
        var pidText = MakeCandidate() with { ProcessId = "0x4D2" };
        var match = ExcelCrashEventReader.Match(Exit, [otherStart, otherApp, pidText]);

        Assert.Equal(ExcelCrashEventStatus.Matched, match.Status);
        var details = Assert.IsType<ExcelCrashEventDetails>(match.Event);
        Assert.Equal(ExitTime.AddSeconds(3), details.TimeUtc);
        Assert.Equal("16.0.20625.20142", details.ApplicationVersion);
        Assert.Equal("EXCEL.EXE", details.ModuleName);
        Assert.Equal("0x000000000045C06F", details.ModuleOffset);
        Assert.Equal("0xC0000005", details.ExceptionCode);
    }

    [Fact]
    public void Match_ConflictingExactMatchesAreAmbiguous()
    {
        var conflicting = MakeCandidate() with { ModuleName = "VBE7.DLL" };
        var result = ExcelCrashEventReader.Match(Exit, [MakeCandidate(), conflicting]);

        Assert.Equal(ExcelCrashEventStatus.Ambiguous, result.Status);
        Assert.Null(result.Event);
    }

    [Fact]
    public void Match_PreservesNoMatchAndRejectsUnconfirmedExits()
    {
        var tooLate = MakeCandidate() with { TimeUtc = ExitTime.AddMinutes(6) };
        var invalidTime = MakeCandidate() with { ProcessCreationTime = "0x1Z" };
        Assert.Equal(ExcelCrashEventStatus.NotFound, ExcelCrashEventReader.Match(Exit, [tooLate, invalidTime]).Status);
        Assert.Equal(ExcelCrashEventStatus.NotApplicable, ExcelCrashEventReader.Match(Exit with { ExitCode = 0 }, []).Status);
        Assert.Equal(ExcelCrashEventStatus.Unavailable, ExcelCrashEventReader.Match(Exit with { ExitedAtUtc = null }, []).Status);
    }

    [Fact]
    public void Match_OmitsPathAndInvalidModuleMetadata()
    {
        var candidate = MakeCandidate() with
        {
            ModuleName = @"C:\private\path\EXCEL.EXE",
            ApplicationVersion = "unexpected private value",
            ModuleOffset = "not-a-number",
            ExceptionCode = "100000000"
        };
        var result = ExcelCrashEventReader.Match(Exit, [candidate]);

        Assert.Equal(ExcelCrashEventStatus.Matched, result.Status);
        var details = Assert.IsType<ExcelCrashEventDetails>(result.Event);
        Assert.Null(details.ModuleName);
        Assert.Null(details.ApplicationVersion);
        Assert.Null(details.ModuleOffset);
        Assert.Null(details.ExceptionCode);
    }

    [Fact]
    public void Match_RecordLimitNeverReportsAUniqueMatchFromAnIncompleteRead()
    {
        var candidates = Enumerable.Repeat(MakeCandidate(), ExcelCrashEventReader.MaximumEvents + 1);
        var result = ExcelCrashEventReader.Match(Exit, candidates);

        Assert.Equal(ExcelCrashEventStatus.Unavailable, result.Status);
        Assert.Null(result.Event);
    }

    [Fact]
    public void Match_RespectsCancellation()
    {
        Assert.Throws<OperationCanceledException>(() =>
            ExcelCrashEventReader.Match(Exit, [MakeCandidate()], new CancellationToken(canceled: true)));
    }

    private static ExcelCrashEventCandidate MakeCandidate(long? start = null) =>
        new(ExitTime.AddSeconds(3), "EXCEL.EXE", "16.0.20625.20142", "EXCEL.EXE",
            "000000000045c06f", "c0000005", Identity.ProcessId, (ulong)(start ?? Identity.StartedAtUtcFileTime));
}
