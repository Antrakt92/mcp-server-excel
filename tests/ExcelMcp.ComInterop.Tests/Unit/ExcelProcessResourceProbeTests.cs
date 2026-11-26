using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelProcessResourceProbeTests
{
    [Fact]
    public void Read_CurrentIdentityReturnsProcessCounters()
    {
        using var process = Process.GetCurrentProcess();
        var identity = new ExcelProcessIdentity(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());

        var result = ExcelProcessResourceProbe.Read(identity);

        // A headless test host can deny one or both GUI counters. Handle and
        // private-byte counters still prove this exact native observation.
        Assert.True(result.Status is ExcelProcessResourceStatus.Observed or ExcelProcessResourceStatus.Partial);
        Assert.True(result.TotalHandles > 0);
        Assert.True(result.PrivateBytes > 0);
        Assert.Equal(result.Status == ExcelProcessResourceStatus.Observed,
            result.GdiObjects is not null && result.UserObjects is not null);
        Assert.True(result.ObservedAtUtc > process.StartTime.ToUniversalTime());
    }

    [Fact]
    public void Read_ReusedProcessNumberCannotExposeCounters()
    {
        using var process = Process.GetCurrentProcess();
        var wrongStart = process.StartTime.ToUniversalTime().ToFileTimeUtc() - 1;

        var result = ExcelProcessResourceProbe.Read(new ExcelProcessIdentity(process.Id, wrongStart));

        Assert.Equal(ExcelProcessResourceStatus.Unavailable, result.Status);
        Assert.Null(result.ObservedAtUtc);
        Assert.Null(result.TotalHandles);
        Assert.Null(result.PrivateBytes);
        Assert.Null(result.GdiObjects);
        Assert.Null(result.UserObjects);
    }
}
