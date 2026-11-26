using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service.Workers;

namespace Sbroenne.ExcelMcp.Service;

internal sealed partial class ExcelController
{
    internal const string ProcessDiagnosticsCommand = "session.process-diagnostics";
    private readonly SemaphoreSlim _processDiagnostics = new(1);

    private async Task<ServiceResponse> ReadProcessDiagnosticsAsync(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
            throw new ArgumentException("A workbook session ID is required for process diagnostics.");
        using var args = JsonDocument.Parse(request.Args ?? "{}");
        var includeCrashEvent = args.RootElement.TryGetProperty("includeCrashEvent", out var option)
            && option.GetBoolean();
        if (!_processDiagnostics.Wait(0, request.CallerCancellation))
            return ExcelWorkerHost.Failure(request, "ServiceBusy", "Another Excel process diagnostic is running.");

        try
        {
            return await Task.Run(() => ReadProcessDiagnostics(request, request.SessionId, includeCrashEvent),
                request.CallerCancellation);
        }
        finally
        {
            _processDiagnostics.Release();
        }
    }

    private ServiceResponse ReadProcessDiagnostics(ServiceRequest request, string sessionId, bool includeCrashEvent)
    {
        var binding = FindOwnedProcess(sessionId);
        if (binding is { } owned)
        {
            var resources = ExcelProcessResourceProbe.Read(owned.Identity);
            if (resources.Status is ExcelProcessResourceStatus.Observed or ExcelProcessResourceStatus.Partial
                && FindOwnedProcess(sessionId) == binding)
                return ProcessDiagnosticResult(request, "active", ExcelMcpService.DescribeExcelProcess(owned.Identity),
                    resources, nativeExit: null, crashEvent: null);
        }

        // Once the live binding is gone, use only exact retained evidence for this session.
        var terminal = _closedSessions.GetValueOrDefault(sessionId);
        if (!_sessions.ContainsKey(sessionId)
            && ((terminal.Owned ? terminal.Excel : null) ?? _closeHistory.FindUniqueOwnedProcess(sessionId))
                is { } identity
            && FindNativeExit(identity) is { State: ExcelNativeProcessState.Exited } exit)
        {
            var crashEvent = includeCrashEvent ? ExcelCrashEventReader.Read(exit, request.CallerCancellation) : null;
            var nativeExit = CloseHistory.NativeDiagnostic(exit);
            nativeExit["exitCodeHex"] = exit.ExitCode?.ToString("X8", CultureInfo.InvariantCulture) is { } hex
                ? "0x" + hex : null;
            nativeExit["exitedAtUtc"] = exit.ExitedAtUtc;
            return ProcessDiagnosticResult(request, "exited", ExcelMcpService.DescribeExcelProcess(identity),
                resources: null, nativeExit, crashEvent);
        }

        return ProcessDiagnosticResult(request, "unknown", ExcelMcpService.DescribeExcelProcess(null),
            resources: null, nativeExit: null, crashEvent: null);
    }

    private bool WasOwnedSession(ExcelWorkerClient worker, string sessionId, ExcelProcessIdentity? identity) =>
        identity is not null && _publishedSnapshots.TryGetValue(worker.Registration.PipeName, out var published)
        && (published.Registration.ReleasedExcel == identity
            || published.Registration is { OwnsExcel: true, ExcelPreserved: false }
                && published.Registration.Excel == identity)
        && published.Sessions.Any(session => session.SessionId == sessionId && !session.IsAttached);

    private static bool HasCurrentOwnedProcess(WorkerSnapshot published, WorkerSnapshot? live) =>
        live is { Releasing: false, Registration: { OwnsExcel: true, ExcelPreserved: false, ReleaseConfirmed: false, Excel: { } identity } }
        && published.Registration.Excel == identity;

    private (ExcelWorkerClient Worker, ExcelProcessIdentity Identity)? FindOwnedProcess(string sessionId)
    {
        lock (_shutdownLock)
        {
            if (!_sessions.TryGetValue(sessionId, out var worker)
                || !_publishedSnapshots.TryGetValue(worker.Registration.PipeName, out var published)
                || worker.Snapshot is not { } live
                || published.Releasing || worker.GetStatusFailure(live, published) is not null
                || !HasCurrentOwnedProcess(published, live) || live.Registration.Excel is not { } identity
                || !published.Sessions.Any(session => session.SessionId == sessionId && !session.IsAttached))
                return null;
            return (worker, identity);
        }
    }

    private static ServiceResponse ProcessDiagnosticResult(ServiceRequest request, string status,
        SessionExcelProcess excelProcess, ExcelProcessResourceSnapshot? resources, JsonObject? nativeExit,
        ExcelCrashEventResult? crashEvent)
    {
        return new ServiceResponse
        {
            Success = true,
            Command = request.Command,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                request.SessionId,
                status,
                excelProcess,
                resources,
                nativeExit,
                crashEvent
            }, ServiceProtocol.JsonOptions)
        };
    }
}
