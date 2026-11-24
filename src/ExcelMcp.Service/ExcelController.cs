using System.Collections.Concurrent;
using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>Host-owned controller. Excel objects and synchronous workbook work live only in private workers.</summary>
internal sealed class ExcelController : IDisposable
{
    private readonly ProductProcessLauncher _launcher;
    private readonly WorkerRegistry _registry;
    private readonly Func<ExcelWorkerClient, Task> _initializeWorker;
    private readonly WorkerProcess _identity = WorkerProcess.Current;
    private readonly string _generation = Guid.NewGuid().ToString("N");
    private readonly DateTime _startTime = DateTime.UtcNow;
    private readonly ConcurrentDictionary<string, ExcelWorkerClient> _workers = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, ExcelWorkerClient> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, WorkerSnapshot> _publishedSnapshots = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, (ExcelWorkerClient Worker, OpenWorkbookInfo Workbook)> _selections = new(StringComparer.Ordinal);
    private sealed record PathReservation(string Token, string? SessionId = null)
    {
        internal (long Revision, bool RequiresLocation)? SaveAsReply { get; set; }
    }
    private readonly WorkbookPathReservations<PathReservation> _paths = new((first, second) =>
        first.SessionId is not null && second.SessionId is not null
            ? first.SessionId == second.SessionId : first.Token == second.Token);
    private readonly SemaphoreSlim _discovery = new(1, 1);
    private readonly ConcurrentDictionary<string, byte> _newWorkers = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, byte> _abandonedWorkers = new(StringComparer.Ordinal);
    private enum SessionEnd { Closed, SavedAndClosed, WorkerExited, ExcelExited }
    private readonly record struct EndedSession(SessionEnd Reason, ExcelProcessIdentity? Excel = null);
    private readonly ConcurrentDictionary<string, EndedSession> _closedSessions = new(StringComparer.Ordinal);
    // Active close outcomes must survive eviction from the bounded diagnostic history.
    private readonly Dictionary<string, ExcelProcessIdentity> _unconfirmedCloseAttempts = new(StringComparer.Ordinal);
    private readonly CloseHistory _closeHistory = new();
    private readonly OperationReceipts _operationReceipts = new();
    private readonly SemaphoreSlim _packageComparisons = new(1, 1);
    private readonly CancellationTokenSource _stop = new();
    private readonly DaemonHost _host;
    private readonly DaemonHost _controlHost;
    private readonly Lock _shutdownLock = new();
    private readonly Lock _recentExcelExitLock = new();
    private readonly LinkedList<ExcelNativeProcessObservation> _recentExcelExits = new();
    private Task<string?>? _shutdown;
    private readonly Task _poll;
    private readonly Lock _utilityLock = new();
    private Task<ExcelWorkerClient>? _utility;
    private bool _disposed;
    internal event Action<ExcelProcessIdentity>? ExcelProcessObserved;
    internal event Action<ExcelProcessIdentity>? ExcelProcessPreserved;
    internal event Action<ExcelProcessIdentity>? ExcelProcessReleased;
    internal Action<string, string>? BeforeRecordPathPublicationHookForTests { get; set; }
    internal Func<ExcelWorkerClient, Task>? BeforeWorkerSnapshotReadForTests { get; set; }
    internal Action<ExcelWorkerClient>? AfterExitedWorkerSnapshotRemovalForTests { get; set; }
    internal Action<string>? BeforeSettledSessionRemovalForTests { get; set; }

    internal ExcelController(
        Func<ServiceRequest, Func<Task<ServiceResponse>>, Task<ServiceResponse>>? trackRequest = null,
        ProductProcessLauncher? launcher = null, WorkerRegistry? registry = null,
        Func<ExcelWorkerClient, Task>? initializeWorker = null)
    {
        _launcher = launcher ?? ProductProcessLauncher.Current();
        _registry = registry ?? new WorkerRegistry();
        _initializeWorker = initializeWorker ?? (worker => worker.InitializeAsync());
        Task<ServiceResponse> ProcessTrackedAsync(ServiceRequest request) => trackRequest is null
            ? ProcessAsync(request) : trackRequest(request, () => ProcessAsync(request));
        _host = new DaemonHost(ProcessTrackedAsync);
        _controlHost = new DaemonHost(request => ServiceClient.IsControlCommand(request.Command)
            ? ProcessTrackedAsync(request)
            : Task.FromResult(ExcelWorkerHost.Failure(request, "InvalidInput", "This connection accepts service controls only.")));
        _poll = PollAsync();
    }

    internal static bool IsWorkerMode(string[] args) => ExcelWorkerHost.IsWorkerMode(args);
    internal static Task<int> RunWorkerAsync(string[] args) => ExcelWorkerHost.RunAsync(args);
    internal Task RunAsync(string pipeName) => Task.WhenAll(
        _host.RunAsync(pipeName), _controlHost.RunAsync(ServiceClient.GetControlPipeName(pipeName)));

    internal IReadOnlyList<SessionDescriptor> Sessions => _sessions
        .SelectMany(entry => _publishedSnapshots.TryGetValue(entry.Value.Registration.PipeName, out var snapshot)
            ? snapshot.Sessions.Where(session => session.SessionId == entry.Key) : [])
        .DistinctBy(session => session.SessionId).ToArray();

    internal IReadOnlyCollection<ExcelProcessIdentity> ExcelProcesses =>
        _workers.Values.Where(worker => worker.Registration.OwnsExcel)
            .Select(worker => worker.Registration.Excel).OfType<ExcelProcessIdentity>().Distinct().ToArray();

    internal async Task<ServiceResponse> ProcessAsync(ServiceRequest request, bool waitForWorkerCompletion = false)
    {
        if (request.OperationId is null)
            return ExcelMcpService.AttachRequestContext(request,
                await ProcessCoreAsync(request, waitForWorkerCompletion).ConfigureAwait(false));

        if (!OperationReceipts.IsValidSessionId(request.SessionId))
            return ExcelWorkerHost.Failure(request, "InvalidInput",
                "operationId requires a session and one of table.append, session.save, or vba.invoke.");
        Guid id;
        try
        {
            id = OperationReceipts.ValidateTracking(request.Command, request.OperationId);
        }
        catch (ArgumentException ex)
        {
            return ExcelWorkerHost.Failure(request, "InvalidInput", ex.Message);
        }
        var admission = _operationReceipts.TryBegin(id, request.Command, request.SessionId!, out var operation);
        if (admission != OperationReceipts.Admission.Accepted)
            return admission == OperationReceipts.Admission.Duplicate
                ? ExcelWorkerHost.Failure(request, "Conflict",
                    "This operationId already has a receipt. Look it up before continuing. No duplicate work started.")
                : ExcelWorkerHost.Failure(request, "ResourceLimit",
                    "The host's operation receipt capacity is full. No workbook work started.");

        var completion = CompleteTrackedRequestAsync(request, operation);
        if (waitForWorkerCompletion)
            return await completion.ConfigureAwait(false);
        try
        {
            return await completion.WaitAsync(request.CallerCancellation).ConfigureAwait(false);
        }
        catch (OperationCanceledException ex) when (request.CallerCancellation.IsCancellationRequested)
        {
            var response = CreateWorkerError(request, ex, sent: true);
            _operationReceipts.ObserveInitialFailure(operation, response);
            return _operationReceipts.IncludeEvidenceInFailure(operation, response);
        }
    }

    private async Task<ServiceResponse> CompleteTrackedRequestAsync(
        ServiceRequest request, OperationReceipts.Operation operation)
    {
        var response = ExcelMcpService.AttachRequestContext(request,
            await ProcessCoreAsync(request, waitForWorkerCompletion: true, operation).ConfigureAwait(false));
        if (request.LateWorkerResponse is { } late)
        {
            _operationReceipts.ObserveInitialFailure(operation, response);
            _ = RecordLateOutcomeAsync(operation, late, response);
            return _operationReceipts.IncludeEvidenceInFailure(operation, response);
        }
        RecordOutcome(operation, response);
        return _operationReceipts.IncludeEvidenceInFailure(operation, response);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Late receipt collection reports unknown if the worker response cannot be recovered. The original command is never repeated.")]
    private async Task RecordLateOutcomeAsync(OperationReceipts.Operation operation, Task<ServiceResponse?> late,
        ServiceResponse initialResponse)
    {
        try
        {
            if (await late.ConfigureAwait(false) is { } response)
                RecordOutcome(operation, response);
            else
                RecordOutcome(operation, initialResponse, forceUnknown: true);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceWarning("Could not recover an Excel operation receipt: {0}", ex.GetType().Name);
            RecordOutcome(operation, initialResponse, forceUnknown: true);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An invalid result payload is reported as an unknown receipt. The original worker reply is preserved and the workbook command is never repeated.")]
    private void RecordOutcome(OperationReceipts.Operation operation, ServiceResponse response, bool forceUnknown = false)
    {
        try
        {
            _operationReceipts.Complete(operation, response, forceUnknown);
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or ArgumentException)
        {
            _operationReceipts.Publish(operation, "unknown");
        }
    }

    internal const string PackageComparisonCommand = "session.compare-packages";

    private async Task<ServiceResponse> ComparePackagesAsync(ServiceRequest request)
    {
        if (request.SessionId is not null)
            throw new ArgumentException("Compare two saved workbook files without a session ID.");

        var options = ServiceRegistry.DeserializeArgs<PackageComparisonArguments>(request.Args);
        string beforePath = FilePathValidation.NormalizeAbsoluteWindowsPath(options.BeforePath ?? "");
        string afterPath = FilePathValidation.NormalizeAbsoluteWindowsPath(options.AfterPath ?? "");
        if (!_packageComparisons.Wait(0, request.CallerCancellation))
            return ExcelWorkerHost.Failure(request, "ServiceBusy", "Another workbook package comparison is running.");

        try
        {
            var result = await Task.Run(() => WorkbookPackageComparison.Compare(
                beforePath, afterPath, options.AllowedParts, request.CallerCancellation, options.IncludeStructure), request.CallerCancellation);
            return new ServiceResponse
            {
                Success = true,
                Command = request.Command,
                Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
            };
        }
        finally
        {
            _packageComparisons.Release();
        }
    }

    private sealed class PackageComparisonArguments
    {
        public string? BeforePath { get; init; }
        public string? AfterPath { get; init; }
        public string[]? AllowedParts { get; init; }
        public bool IncludeStructure { get; init; }
    }

    private static bool StartsWorkbookWorker(string command) =>
        command is "session.open" or "session.create" or "session.test" or "sheet.copy-to-file" or "sheet.move-to-file";

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The controller is a failure-transport boundary. It reports the exact worker/transport cause and uncertain workbook outcome without replaying, saving, or discarding work.")]
    private async Task<ServiceResponse> ProcessCoreAsync(ServiceRequest request, bool waitForWorkerCompletion,
        OperationReceipts.Operation? tracked = null)
    {
        List<string> reservedPaths = [];
        var reservation = new PathReservation(Guid.NewGuid().ToString("N"),
            request.Command == "workbook.save-as" ? request.SessionId : null);
        ExcelWorkerClient? worker = null;
        WorkerRegistration? trackedOwner = null;
        CloseHistory.Target? closeTarget = null;
        JsonObject? priorCloseStatus = null;
        var testTimings = request.Command == "session.test" ? new SessionDiagnostics.TimingCollection() : null;
        var sent = false;
        try
        {
            request.CallerCancellation.ThrowIfCancellationRequested();
            ServiceRegistry.ValidateCommandArguments(request.Command, request.Args);
            if (tracked is not null)
                _operationReceipts.RecordRequestedProcedure(tracked, request.Args);
            if (request.Command.StartsWith("session.", StringComparison.Ordinal))
                ExcelMcpService.ValidateSessionActionArguments(request.Command["session.".Length..], request.Args);
            if (request.Command == "service.ping")
                return new ServiceResponse { Success = true };
            if (request.Command == "service.status")
                return new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new ServiceStatus
                    {
                        Running = true,
                        ProcessId = Environment.ProcessId,
                        SessionCount = _sessions.Count,
                        StartTime = _startTime
                    }, ServiceProtocol.JsonOptions)
                };
            if (request.Command == "session.list")
                return SessionList();
            if (request.Command == OperationReceipts.LookupCommand)
                return _operationReceipts.Lookup(request.Args, request.SessionId);
            if (request.Command == "service.shutdown")
            {
                var report = await ShutdownAsync();
                _host.RequestShutdownAfterResponse();
                _controlHost.RequestShutdownAfterResponse();
                return new ServiceResponse { Success = true, PartialState = report };
            }
            if (_shutdown is not null)
                return ExcelWorkerHost.Failure(request, "SessionUnavailable", "The service is stopping. No new work was started.");
            if (request.Command == PackageComparisonCommand)
                return await ComparePackagesAsync(request);
            if (request.Command == "session.discover")
                return await DiscoverAsync(request.CallerCancellation, waitForWorkerCompletion);

            if (StartsWorkbookWorker(request.Command))
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                string[] operands = request.Command.StartsWith("sheet.", StringComparison.Ordinal)
                    ? ["sourceFile", "targetFile"] : ["filePath"];
                foreach (var operand in operands)
                {
                    var path = arguments.RootElement.TryGetProperty(operand, out var value) ? value.GetString() : null;
                    if (string.IsNullOrWhiteSpace(path))
                        return ExcelWorkerHost.Failure(request, "InvalidInput", $"{operand} is required");
                    var normalized = WorkbookLocation.GetOwnershipKey(path);
                    if (reservedPaths.Contains(normalized, StringComparer.OrdinalIgnoreCase))
                        continue;
                    if (!_paths.TryAdd(normalized, reservation))
                        return ExcelWorkerHost.Failure(request, "Conflict", "This workbook is already reserved by another session or pending request.");
                    reservedPaths.Add(normalized);
                }
                using (testTimings?.BeginStage("controllerWorkerStartup"))
                    worker = await StartWorkerAsync(null, request.CallerCancellation);
            }
            else if (request.Command == "session.attach")
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                var selection = arguments.RootElement.TryGetProperty("workbookId", out var selected)
                    && selected.ValueKind == JsonValueKind.String ? selected.GetString() : null;
                if (selection is null || !_selections.TryGetValue(selection, out var selectedWorkbook))
                    return ExcelWorkerHost.Failure(request, "InvalidInput",
                        selection is null ? "workbookId from session discover is required."
                            : "Unknown open-workbook selection. Discover open workbooks first.");
                worker = selectedWorkbook.Worker;
                if (selectedWorkbook.Workbook.Location is { } location)
                {
                    var keys = selectedWorkbook.Workbook.SessionId is { } selectedSession
                        && _publishedSnapshots.TryGetValue(worker.Registration.PipeName, out var snapshot)
                        && snapshot.OwnershipKeys.TryGetValue(selectedSession, out var ownedPaths)
                        ? ownedPaths.Where(key =>
                            WorkbookLocation.MatchesReportedLocation(key, location)).ToArray()
                        : [WorkbookLocation.GetOwnershipKey(location)];
                    foreach (var key in keys)
                    {
                        if (!_paths.TryGetValue(key, out var existing)
                            || existing.SessionId is null || existing.SessionId != selectedWorkbook.Workbook.SessionId)
                        {
                            if (!_paths.TryAdd(key, reservation))
                                return ExcelWorkerHost.Failure(request, "Conflict", "This workbook is reserved by another session or pending request.");
                            reservedPaths.Add(key);
                        }
                    }
                }
            }
            else if (request.Command.StartsWith("powerquerytrace.", StringComparison.Ordinal)
                && JsonNode.Parse(request.Args ?? "{}")?["workbookSessionId"] is JsonValue workbookFilter
                && workbookFilter.TryGetValue<string>(out var filteredSession)
                && _sessions.TryGetValue(filteredSession, out var traceWorker))
                worker = traceWorker;
            else if (!ExcelMcpService.IsSessionlessCommand(request.Command)
                && request.SessionId is { } sessionId && _sessions.TryGetValue(sessionId, out var owner))
            {
                worker = owner;
                if (tracked is not null)
                    trackedOwner = owner.Registration;
                if (request.Command == "session.close")
                {
                    var registration = owner.Registration;
                    closeTarget = new(sessionId, registration.Excel ?? registration.ReleasedExcel);
                    priorCloseStatus = owner.BuildStatusDiagnostic(_publishedSnapshots.GetValueOrDefault(owner.Registration.PipeName));
                }
            }
            else if (request.Command == "session.close" && request.SessionId is { } closedId
                && _closedSessions.TryGetValue(closedId, out var end) && end.Reason is SessionEnd.Closed or SessionEnd.SavedAndClosed)
                return ExcelMcpService.AlreadyClosedResponse(request, saved: end.Reason == SessionEnd.SavedAndClosed)
                    .WithCloseTimings(new Dictionary<string, double>());
            else if (!ExcelMcpService.IsSessionlessCommand(request.Command) && request.SessionId is { } lostId
                && _closedSessions.TryGetValue(lostId, out var lostEnd) && lostEnd.Reason is SessionEnd.WorkerExited or SessionEnd.ExcelExited)
                return LostSessionResponse(request, lostId, lostEnd);
            else
            {
                lock (_utilityLock)
                {
                    if (_utility is null || _utility.IsFaulted || _utility.IsCanceled
                        || _utility.IsCompletedSuccessfully && (!_utility.Result.Registration.Worker.IsAlive()
                            || _utility.Result.Snapshot?.Releasing == true))
                        _utility = StartWorkerAsync(null, CancellationToken.None);
                }
                worker = await _utility.WaitAsync(request.CallerCancellation);
            }

            if (reservedPaths.Count > 0)
            {
                var associated = reservation with { Token = worker.Registration.PipeName };
                foreach (var path in reservedPaths)
                    _paths.TryUpdate(path, associated, reservation);
                reservation = associated;
            }
            request.CallerCancellation.ThrowIfCancellationRequested();
            if (_shutdown is not null)
                return ExcelWorkerHost.Failure(request, "SessionUnavailable", "The service is stopping. No new work was started.");
            if (request.Command == "workbook.save-as")
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                if (arguments.RootElement.TryGetProperty("targetPath", out var target)
                    && target.ValueKind == JsonValueKind.String && target.GetString() is { } targetPath)
                {
                    var normalized = WorkbookLocation.GetOwnershipKey(FilePathValidation.NormalizeAbsoluteWindowsPath(targetPath));
                    if (!_paths.TryReserveAlias(normalized, reservation,
                        existing => existing.Token == "recorded" && existing.SessionId == request.SessionId,
                        out var added))
                        return ExcelWorkerHost.Failure(request, "Conflict", "The Save As destination belongs to another session or pending request.");
                    if (added)
                        reservedPaths.Add(normalized);
                }
            }
            if (!worker.Registration.Worker.IsAlive())
                throw new IOException("The Excel worker exited before this request was sent. Discover open workbooks to reconnect; no work was started.");
            if (tracked is not null && trackedOwner is not null)
            {
                lock (_recentExcelExitLock)
                    _operationReceipts.BindWorker(tracked, trackedOwner,
                        _recentExcelExits.FirstOrDefault(exit => exit.Identity == trackedOwner.Excel));
            }
            if (closeTarget is not null)
                RememberCloseDispatch(closeTarget, worker);
            sent = true;
            ServiceResponse response;
            using (testTimings?.BeginStage("controllerWorkerRequest"))
            {
                var completion = CompleteWorkerRequestAsync(request, worker, reservedPaths, reservation, closeTarget, priorCloseStatus,
                    tracked?.DispatchId);
                response = waitForWorkerCompletion
                    ? await completion
                    : await completion.WaitAsync(request.CallerCancellation);
            }
            return testTimings is not null && response.Success
                ? response.WithResultTimings("testTimingsMs", testTimings.Snapshot())
                : response;
        }
        catch (Exception ex)
        {
            if (worker is not null && StartsWorkbookWorker(request.Command))
                _abandonedWorkers[worker.Registration.PipeName] = 0;
            var failure = CreateWorkerError(request, ex, sent);
            if (!sent && closeTarget is not null)
                RecordCloseAttempt(closeTarget, failure, priorCloseStatus);
            return failure;
        }
        finally
        {
            if (!sent)
                foreach (var path in reservedPaths)
                    _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The controller owns late worker completion after caller cancellation. It records exact results and confirmed state, and reports/logs transport failures rather than losing a faulted background task or retrying edits.")]
    private async Task<ServiceResponse> CompleteWorkerRequestAsync(
        ServiceRequest request, ExcelWorkerClient worker, List<string> reservedPaths, PathReservation reservation,
        CloseHistory.Target? closeTarget = null, JsonObject? priorCloseStatus = null, Guid? dispatchId = null)
    {
        ServiceResponse? completedResponse = null;
        ExcelOperationBlocker? priorSaveBlocker = null;
        try
        {
            var execution = await worker.ExecuteWithAdmissionAsync(request, request.CallerCancellation,
                dispatchId, dispatchId is null ? null : _operationReceipts.ObserveWorkerEvidence);
            var response = execution.Response;
            completedResponse = response;
            priorSaveBlocker = execution.PriorSaveBlocker;
            if (closeTarget is not null)
                RecordCloseAttempt(closeTarget, response, priorCloseStatus, priorSaveBlocker, accepted: execution.Accepted);
            if (!execution.Accepted)
                foreach (var path in reservedPaths)
                    _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
            if (execution.Accepted && request.Command == "workbook.save-as" && reservedPaths.Count > 0
                && CanReconcileSaveAsClaim(response))
            {
                // A dialog, pending native work, or a failed status capture can defer reconciliation.
                // Wait for a settled snapshot at or after this reply before dropping the claim.
                // A native wait may have given an untitled workbook a path before Excel could report it.
                lock (_shutdownLock)
                    reservation.SaveAsReply = (execution.Snapshot.Revision,
                        response.PartialState is not null && response.ErrorCategory is "Timeout" or "Cancelled");
            }
            var applied = ApplyWorkerSnapshot(worker);
            lock (_shutdownLock)
            {
                if (!_workers.TryGetValue(worker.Registration.PipeName, out var current) || current != worker)
                    return response;
                var published = _publishedSnapshots.GetValueOrDefault(worker.Registration.PipeName);
                var live = worker.Snapshot;
                if (response.SessionId is { } created && published is { } createdSnapshot
                    && createdSnapshot.Sessions.Any(session => session.SessionId == created))
                {
                    if (response.Success && request.Command == "session.attach"
                        && JsonNode.Parse(request.Args!)?["workbookId"]?.GetValue<string>() is { } selectionId
                        && _selections.TryGetValue(selectionId, out var selection) && selection.Worker == worker)
                        _selections[selectionId] = (worker, selection.Workbook with { SessionId = created });
                    var currentPaths = RecordedPaths(createdSnapshot, created);
                    foreach (var path in reservedPaths)
                        if (currentPaths.Contains(path))
                            RecordPath(worker, path, created);
                    _sessions[created] = worker;
                }
                if (response.Success && request.SessionId is { } reconciled
                    && worker.StatusError is null && live is { StatusError: null } && !HasPendingWork(live)
                    && published is { StatusError: null } settledSnapshot
                    && settledSnapshot.Revision == live.Revision
                    && !HasPendingWork(settledSnapshot)
                    && settledSnapshot.Sessions.FirstOrDefault(session => session.SessionId == reconciled)?.FilePath is not null)
                {
                    var currentPaths = RecordedPaths(settledSnapshot, reconciled);
                    foreach (var path in currentPaths)
                        RecordPath(worker, path, reconciled);
                    foreach (var old in _paths.Where(path => path.Value.Token == "recorded" && path.Value.SessionId == reconciled
                        && !currentPaths.Contains(path.Key)))
                        _paths.TryRemove(old);
                }
                if (response.Success && request.Command == "session.close" && request.SessionId is { } closed)
                {
                    foreach (var path in _paths.Where(path => path.Value.SessionId == closed))
                        _paths.TryRemove(path.Key, out _);
                    _sessions.TryRemove(closed, out _);
                }
                if (worker.StatusError is null && worker.Snapshot is { StatusError: null } snapshot
                    && !HasPendingWork(snapshot) && (response.Success || response.PartialState is null))
                {
                    var retainedPaths = snapshot.Sessions.SelectMany(session => RecordedPaths(snapshot, session.SessionId))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    foreach (var path in reservedPaths)
                        if (!retainedPaths.Contains(path))
                            _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
                }
            }
            if (applied is { Registration.ReleaseConfirmed: true })
                RemoveReleasedWorker(worker);
            else if (request.Command is "session.close" or "session.test" or "sheet.copy-to-file" or "sheet.move-to-file"
                || !response.Success && reservedPaths.Count > 0)
                await ReleaseEmptyWorkerAsync(worker);
            return response;
        }
        catch (Exception ex)
        {
            if (request.OperationId is not null)
                CaptureRetainedOperationEvidence(worker.Registration);
            if (StartsWorkbookWorker(request.Command))
                _abandonedWorkers[worker.Registration.PipeName] = 0;
            if (ex is OperationCanceledException && reservedPaths.Count > 0)
            {
                await RefreshWorkerAsync(worker, CancellationToken.None);
                lock (_shutdownLock)
                {
                    if (worker.StatusError is null && worker.Snapshot is { StatusError: null } snapshot && !HasPendingWork(snapshot))
                    {
                        var retainedPaths = snapshot.Sessions.SelectMany(session => RecordedPaths(snapshot, session.SessionId))
                            .ToHashSet(StringComparer.OrdinalIgnoreCase);
                        foreach (var path in reservedPaths)
                            if (!retainedPaths.Contains(path))
                                _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
                    }
                }
            }
            System.Diagnostics.Trace.TraceError("Excel worker request failed: {0}", ex);
            var failure = CreateWorkerError(request, ex, sent: true, completed: completedResponse);
            if (closeTarget is not null)
                RecordCloseAttempt(closeTarget, failure, priorCloseStatus, priorSaveBlocker);
            return failure;
        }
    }

    private void RememberCloseDispatch(CloseHistory.Target target, ExcelWorkerClient worker)
    {
        if (target.Excel is not { } identity)
            return;

        lock (_shutdownLock)
            if (_sessions.TryGetValue(target.SessionId, out var owner) && owner == worker
                && !_closedSessions.ContainsKey(target.SessionId))
                _unconfirmedCloseAttempts[target.SessionId] = identity;
    }

    private void RecordCloseAttempt(CloseHistory.Target target, ServiceResponse response, JsonObject? priorStatus,
        ExcelOperationBlocker? priorSaveBlocker = null, bool accepted = false)
    {
        lock (_shutdownLock)
        {
            var confirmation = _closeHistory.Record(target, response, priorStatus, priorSaveBlocker: priorSaveBlocker);
            if (accepted && confirmation.CloseConfirmed)
            {
                // A completed reply can confirm the close independently of worker status publication.
                var confirmedEnd = new EndedSession(confirmation.SaveConfirmed ? SessionEnd.SavedAndClosed : SessionEnd.Closed);
                _closedSessions.AddOrUpdate(target.SessionId, confirmedEnd, (_, previous) =>
                    previous.Reason == SessionEnd.SavedAndClosed ? previous : confirmedEnd);
                _unconfirmedCloseAttempts.Remove(target.SessionId);
                return;
            }

            // A status poll can infer a normal exit before the failed close reply arrives.
            if (!confirmation.CloseConfirmed && target.Excel is { } identity
                && _closedSessions.TryGetValue(target.SessionId, out var end)
                && end is { Reason: SessionEnd.Closed } && end.Excel == identity
                && _closeHistory.HasUnconfirmedFailure(target))
                _closedSessions.TryUpdate(target.SessionId, new(SessionEnd.ExcelExited, identity), end);
        }
    }

    private static ServiceResponse CreateWorkerError(
        ServiceRequest request, Exception ex, bool sent, ServiceResponse? completed = null, bool nativeExitObserved = false) =>
        completed?.WithCompletionFailure(request, ex) ?? new()
        {
            Success = false,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = ex switch
            {
                OperationCanceledException => "Cancelled",
                TimeoutException => "Timeout",
                ArgumentException or JsonException => "InvalidInput",
                InvalidDataException or EndOfStreamException when request.Command == PackageComparisonCommand => "InvalidInput",
                FileNotFoundException or DirectoryNotFoundException when request.Command == PackageComparisonCommand => "NotFound",
                UnauthorizedAccessException when request.Command == PackageComparisonCommand => "Permissions",
                IOException when request.Command == PackageComparisonCommand => "OperationFailed",
                _ => "ServiceUnavailable"
            },
            ExceptionType = ex.GetType().Name,
            ErrorMessage = $"{ex.GetType().Name}: {ex.Message}",
            PartialState = !sent || ex is ArgumentException or JsonException ? null : JsonSerializer.Serialize(new
            {
                outcome = "unknown",
                message = nativeExitObserved
                    ? "Excel exited before the outcome could be confirmed. Check the saved workbook before repeating edits."
                    : "The worker response could not be confirmed. Work may still be running; do not repeat edits automatically."
            }, ServiceProtocol.JsonOptions)
        };

    private ServiceResponse LostSessionResponse(ServiceRequest request, string sessionId, EndedSession session)
    {
        var exit = session.Excel is { } excel ? FindNativeExit(excel) : null;
        var excelExited = session.Reason == SessionEnd.ExcelExited || exit is not null;
        if (exit is null && excelExited && session.Excel is { } identity)
            exit = new(identity, ExcelNativeProcessState.Exited, null, null);

        var response = CreateWorkerError(request,
            new IOException(excelExited
                ? "The workbook's Excel process exited. Its last outcome could not be confirmed; no work was repeated."
                : "The workbook's Excel helper exited. Its last outcome could not be confirmed; no work was repeated."),
            sent: true, nativeExitObserved: excelExited);
        if (request.Command == "session.close")
            _closeHistory.Record(new(sessionId, session.Excel), response, onlyIfMissing: true);
        return exit is null ? response : response.WithPartialState(
            ExcelMcpService.MergeNativeProcessDiagnostics(response.PartialState, exit)!);
    }

    private static bool HasPendingWork(WorkerSnapshot snapshot) =>
        snapshot.PendingOperation is not null
        || JsonNode.Parse(snapshot.SessionList.Result!)!["sessions"]!.AsArray()
            .Any(session => session?["operationPending"]?.GetValue<bool>() == true);

    private static bool CanReconcileSaveAsClaim(ServiceResponse response)
    {
        if (response.Success || response.PartialState is null || response.ErrorCategory is "Timeout" or "Cancelled")
            return true;

        // Input or readiness checks can fail before the new callback starts. Only
        // legacy wait diagnostics from a finished callback may accompany that failure.
        return response.ErrorCategory is "Busy" or "InvalidInput"
            && JsonNode.Parse(response.PartialState) is JsonObject { Count: 1 } state
            && state["operation"] is JsonObject operation
            && operation["waitExpired"] is JsonValue expired && expired.TryGetValue<bool>(out var wasExpired) && wasExpired
            && operation["executionPending"] is JsonValue pending && pending.TryGetValue<bool>(out var isPending) && !isPending;
    }

    private static HashSet<string> RecordedPaths(WorkerSnapshot snapshot, string sessionId) =>
        snapshot.OwnershipKeys[sessionId].ToHashSet(StringComparer.OrdinalIgnoreCase);

    private void RecordPath(ExcelWorkerClient worker, string path, string sessionId)
    {
        BeforeRecordPathPublicationHookForTests?.Invoke(sessionId, path);
        var recorded = new PathReservation("recorded", sessionId);
        if (!_paths.TryPublish(path, recorded,
            existing => existing.SessionId is null && existing.Token == worker.Registration.PipeName))
            throw new InvalidOperationException("The workbook's physical path is reserved by another worker or session.");
    }

    internal async Task<ExcelWorkerClient> StartWorkerAsync(ExcelProcessIdentity? excel, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ExcelWorkerClient worker;
        lock (_shutdownLock)
        {
            if (_shutdown is not null)
                throw new InvalidOperationException("The service is stopping. No Excel worker was started.");
            if (_workers.Count >= 32)
                throw new InvalidOperationException("The service has reached its Excel worker limit. Close unused sessions before retrying.");
            worker = ExcelWorkerClient.Launch(_launcher, _identity, _generation, excel);
            _workers[worker.Registration.PipeName] = worker;
            _newWorkers[worker.Registration.PipeName] = 0;
        }
        try
        {
            await _initializeWorker(worker);
        }
        catch
        {
            _abandonedWorkers[worker.Registration.PipeName] = 0;
            if (!worker.Registration.Worker.IsAlive())
                RetireExitedWorker(worker.Registration);
            throw;
        }
        return worker;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An unused worker is released only after it reports no sessions and no pending work. Release failures retain its registration and appear as a status error, without changing the caller's workbook result.")]
    private async Task ReleaseEmptyWorkerAsync(ExcelWorkerClient worker)
    {
        if (worker.Snapshot is not { Sessions.Count: 0, PendingOperation: null, StatusError: null }
            || _sessions.Values.Contains(worker))
            return;
        try
        {
            var released = await worker.ReleaseAsync(CancellationToken.None);
            var applied = ApplyWorkerSnapshot(worker);
            if (!released.Success)
                throw new InvalidOperationException(released.ErrorMessage);
            if (released.PartialState is not null)
                throw new InvalidOperationException($"The Excel worker did not confirm complete release: {released.PartialState}");
            if (applied is { Registration.ReleaseConfirmed: true })
                RemoveReleasedWorker(worker);
        }
        catch (Exception ex)
        {
            worker.ReportStatusFailure("workerRelease", $"The unused Excel worker could not be released: {ex.Message}", ex);
            System.Diagnostics.Trace.TraceError("Could not release unused Excel worker: {0}", ex);
        }
    }

    private void RemoveReleasedWorker(ExcelWorkerClient worker)
    {
        lock (_shutdownLock)
        {
            var pipeName = worker.Registration.PipeName;
            if (!_workers.TryGetValue(pipeName, out var current) || current != worker)
                return;
            _registry.Remove(pipeName);
            _workers.TryRemove(pipeName, out _);
            _publishedSnapshots.TryRemove(pipeName, out _);
            _newWorkers.TryRemove(pipeName, out _);
            _abandonedWorkers.TryRemove(pipeName, out _);
            foreach (var path in _paths.Where(entry => entry.Value.Token == pipeName
                && entry.Value.SessionId is null))
                _paths.TryRemove(path);
            foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker))
                _selections.TryRemove(selection);
        }
        ClearUtilityWorker(worker);
    }

    private void ClearUtilityWorker(ExcelWorkerClient worker)
    {
        lock (_utilityLock)
            if (_utility is { IsCompletedSuccessfully: true } && _utility.Result == worker)
                _utility = null;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Recorded-status failures are exposed on session status and logged. They never replace a completed operation result or manufacture an empty healthy session list.")]
    private async Task<WorkerSnapshot?> RefreshWorkerAsync(ExcelWorkerClient worker, CancellationToken cancellationToken)
    {
        try
        {
            var recorded = !worker.Registration.Worker.IsAlive()
                ? _registry.Read(worker.Registration.PipeName)
                : null;
            if (recorded is null || !RecoverRecordedWorkerStatus(worker, recorded))
            {
                if (BeforeWorkerSnapshotReadForTests is { } beforeRead)
                    await beforeRead(worker);
                await worker.ReadSnapshotAsync(cancellationToken, _operationReceipts.ObserveWorkerEvidence);
            }
            return ApplyWorkerSnapshot(worker);
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            worker.ReportStatusFailure("snapshotTransport",
                $"Excel worker status could not be refreshed: {ex.GetType().Name}: {ex.Message}", ex);
            System.Diagnostics.Trace.TraceError("Excel worker status could not be refreshed: {0}", ex);
            return null;
        }
    }

    private WorkerSnapshot? ApplyWorkerSnapshot(ExcelWorkerClient worker)
    {
        lock (_shutdownLock)
        {
            if (!_workers.TryGetValue(worker.Registration.PipeName, out var current) || current != worker
                || worker.Snapshot is not { } snapshot)
                return null;
            PublishProcessIdentity(snapshot.Registration);
            var sessions = snapshot.Sessions.Where(session => !_closedSessions.ContainsKey(session.SessionId)).ToArray();
            var settled = !HasPendingWork(snapshot) && snapshot.StatusError is null;
            // Keep the previous published metadata and its claims until every new path is owned.
            // A failed registration deliberately retains any partial claims for safe reconciliation.
            foreach (var session in sessions)
                if (!string.IsNullOrWhiteSpace(session.FilePath))
                    foreach (var ownedPath in RecordedPaths(snapshot, session.SessionId))
                        RecordPath(worker, ownedPath, session.SessionId);

            foreach (var session in sessions)
            {
                if (settled)
                {
                    var currentPaths = RecordedPaths(snapshot, session.SessionId);
                    var hasCurrentPath = !string.IsNullOrWhiteSpace(session.FilePath);
                    if (hasCurrentPath)
                        foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker
                            && entry.Value.Workbook.SessionId == session.SessionId
                            && !currentPaths.Any(key => WorkbookLocation.MatchesReportedLocation(key,
                                entry.Value.Workbook.Location))))
                            _selections[selection.Key] = (worker, selection.Value.Workbook with { Location = session.FilePath });
                    foreach (var old in _paths.Where(entry => entry.Value.SessionId == session.SessionId
                        && (hasCurrentPath && entry.Value.Token == "recorded"
                            || entry.Value.SaveAsReply is { } reply && snapshot.Revision >= reply.Revision
                                && (hasCurrentPath || !reply.RequiresLocation))
                        && !currentPaths.Contains(entry.Key)))
                        _paths.TryRemove(old);
                }
            }
            if (settled)
            {
                EndedSession? observedEnd = null;
                foreach (var removed in _sessions.Where(entry => entry.Value == worker
                    && !snapshot.Sessions.Any(session => session.SessionId == entry.Key)))
                {
                    if (!_closedSessions.ContainsKey(removed.Key))
                    {
                        var end = observedEnd ??= InferSettledSessionEnd(snapshot.Registration);
                        if (end is { Reason: SessionEnd.Closed, Excel: { } identity }
                            && (_unconfirmedCloseAttempts.TryGetValue(removed.Key, out var dispatched) && dispatched == identity
                                || _closeHistory.HasUnconfirmedFailure(new(removed.Key, identity))))
                            end = new(SessionEnd.ExcelExited, identity);
                        _closedSessions.TryAdd(removed.Key, end);
                    }
                    _unconfirmedCloseAttempts.Remove(removed.Key);
                    foreach (var path in _paths.Where(entry => entry.Value.SessionId == removed.Key))
                        _paths.TryRemove(path);
                    BeforeSettledSessionRemovalForTests?.Invoke(removed.Key);
                    _sessions.TryRemove(removed);
                }
            }
            // Once a settled row disappears, a targeted retry must see its
            // terminal outcome and its claims must already be released.
            _publishedSnapshots[worker.Registration.PipeName] = snapshot;
            if (_newWorkers.ContainsKey(snapshot.Registration.PipeName))
                foreach (var session in sessions)
                    _sessions[session.SessionId] = worker;
            return snapshot;
        }
    }

    private EndedSession InferSettledSessionEnd(WorkerRegistration registration)
    {
        if ((registration.Excel ?? registration.ReleasedExcel) is not { } identity)
            return new(SessionEnd.Closed);

        var exit = FindNativeExit(identity);
        if (registration.LastExcelObservation is { } recorded && recorded.Identity == identity
            && CloseHistory.IsStrongerNativeExit(recorded, exit))
            exit = recorded;

        // The worker can remove the session before its exit observation reaches the registry.
        // An exact native probe proves exit, but cannot confirm a workbook close or save.
        if (exit is null && OwnedProcessGuard.TryConfirmExited(identity))
            exit = new(identity, ExcelNativeProcessState.Exited, null, null);
        if (exit is null)
            return new(SessionEnd.Closed);

        RememberNativeExit(exit);
        return new(exit.ExitCode == 0 ? SessionEnd.Closed : SessionEnd.ExcelExited, identity);
    }

    private ExcelNativeProcessObservation? FindNativeExit(ExcelProcessIdentity identity)
    {
        lock (_recentExcelExitLock)
            return _recentExcelExits.FirstOrDefault(observation => observation.Identity == identity);
    }

    private bool RecoverRecordedWorkerStatus(ExcelWorkerClient worker, WorkerRegistration recorded)
    {
        var current = worker.Registration;
        if (!KnownRegistration(current, recorded))
            throw new InvalidOperationException("The retained Excel worker record has a mismatched process identity or build.");

        _operationReceipts.ObserveWorkerEvidence(recorded);

        // An exit observation alone does not restore ownership or acknowledge release.
        if (recorded.LastExcelObservation is { State: ExcelNativeProcessState.Exited } observation)
        {
            if ((current.Excel != observation.Identity && current.ReleasedExcel != observation.Identity)
                || (recorded.Excel != observation.Identity && recorded.ReleasedExcel != observation.Identity))
                throw new InvalidOperationException("The retained native observation does not match the known Excel process identity.");
            RememberRegisteredNativeExit(recorded);
        }
        if (!recorded.ReleaseConfirmed)
            return false;

        worker.RecordSnapshot(new(recorded, [],
            new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
            null, null, new Dictionary<string, IReadOnlyList<string>>(),
            Releasing: true, Revision: (worker.Snapshot?.Revision ?? 0) + 1));
        return true;
    }

    private bool KnownRegistration(WorkerRegistration current, WorkerRegistration recorded) =>
        recorded.PipeName == current.PipeName && recorded.Worker == current.Worker
        && recorded.Controller == _identity && recorded.Generation == _generation
        && recorded.Build == WorkerRegistry.Build;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Retained diagnostics are supplemental; failure to read them cannot replace an operation or transport failure.")]
    private void CaptureRetainedOperationEvidence(WorkerRegistration current)
    {
        try
        {
            if (_registry.Read(current.PipeName) is not { } recorded || !KnownRegistration(current, recorded))
                return;
            _operationReceipts.ObserveWorkerEvidence(recorded);
            if (recorded.LastExcelObservation is { State: ExcelNativeProcessState.Exited } exit
                && (current.Excel == exit.Identity || current.ReleasedExcel == exit.Identity))
                RememberRegisteredNativeExit(recorded);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException)
        {
            System.Diagnostics.Trace.TraceWarning("Could not read retained operation evidence: {0}", ex.GetType().Name);
        }
    }

    private void PublishProcessIdentity(WorkerRegistration registration)
    {
        _operationReceipts.ObserveWorkerEvidence(registration);
        RememberRegisteredNativeExit(registration);
        if (registration.ReleasedExcel is { } finished)
            ExcelProcessReleased?.Invoke(finished);
        if (registration is { OwnsExcel: true, Excel: { } excel })
        {
            ExcelProcessObserved?.Invoke(excel);
            if (registration.ExcelPreserved)
                ExcelProcessPreserved?.Invoke(excel);
        }
    }

    private void RememberRegisteredNativeExit(WorkerRegistration registration)
    {
        if (registration.LastExcelObservation is { State: ExcelNativeProcessState.Exited } observation
            && (registration.Excel == observation.Identity || registration.ReleasedExcel == observation.Identity))
            RememberNativeExit(observation);
    }

    private void RememberNativeExit(ExcelNativeProcessObservation observation)
    {
        lock (_recentExcelExitLock)
        {
            _operationReceipts.ObserveNativeExit(observation);
            var recorded = _recentExcelExits.First;
            while (recorded is not null && recorded.Value.Identity != observation.Identity)
                recorded = recorded.Next;
            if (recorded is null)
            {
                _recentExcelExits.AddLast(observation);
                if (_recentExcelExits.Count > 16)
                    _recentExcelExits.RemoveFirst();
            }
            else if (CloseHistory.IsStrongerNativeExit(observation, recorded.Value))
            {
                recorded.Value = observation;
            }
        }
    }

    private (JsonArray ExcelExits, JsonArray CloseAttempts) RecentCloseEvidence(HashSet<string> listedSessions)
    {
        lock (_shutdownLock)
        {
            lock (_recentExcelExitLock)
            {
                return (
                    new JsonArray(_recentExcelExits.Select(exit => (JsonNode)CloseHistory.NativeDiagnostic(exit)).ToArray()),
                    _closeHistory.Snapshot(_recentExcelExits,
                        session => !listedSessions.Contains(session) && _closedSessions.ContainsKey(session) && !_sessions.ContainsKey(session)
                            && !_paths.Any(path => path.Value.SessionId == session)));
            }
        }
    }

    private ServiceResponse SessionList()
    {
        var sessions = new JsonArray();
        var listedSessions = new HashSet<string>(StringComparer.Ordinal);
        List<string> unavailableWorkers = [];
        foreach (var worker in _workers.Values)
        {
            var live = worker.Snapshot;
            _publishedSnapshots.TryGetValue(worker.Registration.PipeName, out var snapshot);
            var statusFailure = worker.GetStatusFailure(live, snapshot);
            var statusError = statusFailure?.Message;
            if (snapshot is null)
            {
                if (statusError is not null)
                    unavailableWorkers.Add(statusError);
                continue;
            }
            var visibleSessionsBefore = sessions.Count;
            var pending = snapshot.PendingOperation is not null || live is not null && HasPendingWork(live);
            var releasing = snapshot.Releasing || live?.Releasing == true;
            var list = JsonNode.Parse(snapshot.SessionList.Result!)!.AsObject();
            foreach (var session in list["sessions"]!.AsArray())
            {
                if (session?["sessionId"]?.GetValue<string>() is not { } sessionId || !_sessions.ContainsKey(sessionId))
                    continue;
                var copy = session.DeepClone();
                if (pending)
                {
                    copy["operationPending"] = true;
                    copy["canClose"] = false;
                    copy["closeReadiness"] = "blocked";
                }
                if (statusError is not null || releasing)
                {
                    bool blocked = releasing || copy["operationPending"]?.GetValue<bool>() == true
                        || copy["canClose"]?.GetValue<bool>() != true;
                    if (blocked && copy["blockingReason"] is { } priorBlocker)
                        copy["lastKnownBlockingReason"] = priorBlocker.DeepClone();
                    copy["canClose"] = false;
                    copy["blockingReason"] = "WorkerUnavailable";
                    copy["blockingMessage"] = statusError ?? "The Excel worker is being released.";
                    copy["closeReadiness"] = blocked ? "blocked" : "unconfirmed";
                    if (statusFailure?.Diagnostic is { } diagnostic)
                        copy["statusDiagnostic"] = diagnostic.ToJson(live?.Registration ?? snapshot.Registration);
                }
                sessions.Add(copy);
                listedSessions.Add(sessionId);
            }
            if (statusError is not null && sessions.Count == visibleSessionsBefore)
                unavailableWorkers.Add(statusError);
        }
        var (recentExcelExits, recentCloseAttempts) = RecentCloseEvidence(listedSessions);
        if (unavailableWorkers.Count > 0)
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = string.Join(Environment.NewLine, unavailableWorkers),
                PartialState = JsonSerializer.Serialize(new { complete = false, sessions, count = sessions.Count, recentExcelExits, recentCloseAttempts },
                    ServiceProtocol.JsonOptions)
            };
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count, recentExcelExits, recentCloseAttempts }, ServiceProtocol.JsonOptions)
        };
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Discovery reports each failed application's exact inspection/connection error alongside other discovered workbooks. A failed application is never presented as an empty successful inspection.")]
    private async Task<ServiceResponse> DiscoverAsync(CancellationToken cancellationToken, bool waitForWorkerCompletion)
    {
        await _discovery.WaitAsync(cancellationToken);
        try
        {
            var native = RunningExcelWindows.Find();
            var retained = _registry.Read().Concat(_workers.Values.Select(worker => worker.Registration))
                .DistinctBy(record => record.PipeName).ToArray();
            var workbooks = new List<OpenWorkbookInfo>();
            var errors = new List<WorkbookDiscoveryError>(native.Errors);
            var allWorkersComplete = true;
            foreach (var window in native.Windows.DistinctBy(window => window.Process))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var worker = _workers.Values.FirstOrDefault(worker =>
                        worker.Registration.Excel == window.Process && worker.Registration.Worker.IsAlive());
                    if (worker is null)
                    {
                        var record = retained.Where(record => record.Excel == window.Process)
                            .OrderByDescending(record => record.Worker.IsAlive()).FirstOrDefault();
                        if (record is not null)
                        {
                            if (!record.Worker.IsAlive())
                            {
                                errors.Add(new(window.Process.ProcessId,
                                    "The previous Excel worker exited without confirming its outcome. Its last work may have completed. This explicit discovery reconnects for inspection only; no edit was repeated."));
                                var exitedRecords = retained.Where(item =>
                                    item.Excel == window.Process && !item.Worker.IsAlive()).ToArray();
                                foreach (var exited in exitedRecords)
                                    RetireExitedWorker(exited, explicitDiscovery: true);
                                worker = await StartWorkerAsync(window.Process, cancellationToken);
                                foreach (var exited in exitedRecords)
                                    _registry.Remove(exited.PipeName);
                            }
                            else if (record.Build != WorkerRegistry.Build)
                            {
                                errors.Add(new(window.Process.ProcessId,
                                    "Discovery skipped this Excel instance because its worker uses a different ExcelMcp build. " +
                                    "Use the connection that owns the worker if available. " +
                                    "For workbooks already managed by this controller, continue using their existing session IDs."));
                                continue;
                            }
                            else if (record.Controller != _identity && record.Controller.IsAlive())
                            {
                                errors.Add(new(window.Process.ProcessId,
                                    "Discovery skipped this Excel instance because it belongs to another running ExcelMcp controller. " +
                                    "Reuse the owning CLI daemon or MCP connection if available. " +
                                    "For workbooks already managed by this controller, continue using their existing session IDs."));
                                continue;
                            }
                            else
                                worker = new ExcelWorkerClient(record, _identity, _generation);
                        }
                        else
                            worker = await StartWorkerAsync(window.Process, cancellationToken);
                    }
                    if (worker.Snapshot is null)
                        await ClaimRetainedWorkerAsync(worker, cancellationToken);
                    if (worker.Snapshot?.Releasing == true)
                    {
                        errors.Add(new(window.Process.ProcessId, "The previous worker is still releasing Excel. No second Excel owner was started."));
                        continue;
                    }
                    var previousSelections = _selections.Where(entry => entry.Value.Worker == worker)
                        .ToDictionary(entry => entry.Key, entry => entry.Value.Workbook, StringComparer.Ordinal);
                    var completion = CompleteWorkerRequestAsync(new ServiceRequest
                    {
                        Command = "session.discover",
                        CallerCancellation = cancellationToken
                    }, worker, [], new PathReservation("discovery"));
                    var response = waitForWorkerCompletion ? await completion : await completion.WaitAsync(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!response.Success)
                    {
                        errors.Add(new(window.Process.ProcessId, response.ErrorMessage ?? "Excel inspection failed."));
                        continue;
                    }
                    using var result = JsonDocument.Parse(response.Result!);
                    var discovered = result.RootElement.GetProperty("workbooks").Deserialize<List<OpenWorkbookInfo>>(ServiceProtocol.JsonOptions)!;
                    var inspectionErrors = result.RootElement.GetProperty("errors")
                        .Deserialize<List<WorkbookDiscoveryError>>(ServiceProtocol.JsonOptions)!;
                    var workerComplete = result.RootElement.GetProperty("complete").GetBoolean() && inspectionErrors.Count == 0;
                    allWorkersComplete &= workerComplete;
                    PublishDiscoverySelections(worker, previousSelections, discovered, workerComplete, workbooks);
                    errors.AddRange(inspectionErrors);
                }
                catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
                {
                    errors.Add(new(window.Process.ProcessId, $"{ex.GetType().Name}: {ex.Message}"));
                }
            }
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, workbooks, errors, complete = allWorkersComplete && errors.Count == 0 },
                    ServiceProtocol.JsonOptions)
            };
        }
        finally
        {
            _discovery.Release();
        }
    }

    private void PublishDiscoverySelections(ExcelWorkerClient worker,
        IReadOnlyDictionary<string, OpenWorkbookInfo> previous, List<OpenWorkbookInfo> discovered,
        bool complete, List<OpenWorkbookInfo> publicWorkbooks)
    {
        lock (_shutdownLock)
        {
            if (!_workers.TryGetValue(worker.Registration.PipeName, out var currentWorker) || currentWorker != worker
                || worker.Snapshot?.Releasing == true)
                throw new InvalidOperationException("The Excel worker changed during discovery. Discover open workbooks again.");

            if (complete)
            {
                var liveIds = discovered.Select(workbook => workbook.WorkbookId).ToHashSet(StringComparer.Ordinal);
                foreach (var candidate in previous)
                    if (!liveIds.Contains(candidate.Key))
                        _selections.TryRemove(new(candidate.Key, (worker, candidate.Value)));
            }

            foreach (var workbook in discovered)
            {
                var selected = workbook;
                if (_selections.TryGetValue(workbook.WorkbookId, out var current)
                    && !ReferenceEquals(current.Workbook, previous.GetValueOrDefault(workbook.WorkbookId)))
                {
                    if (current.Worker != worker)
                        throw new InvalidOperationException("Two Excel workers reported the same workbook selection.");
                    selected = current.Workbook;
                }
                else
                    _selections[workbook.WorkbookId] = (worker, workbook);

                // Keep the retained worker's ID privately for path ownership during attach.
                // Other commands can use the ID only after this controller publishes it.
                if (selected.SessionId is { } sessionId
                    && (!_sessions.TryGetValue(sessionId, out var owner) || owner != worker))
                    selected = selected with { SessionId = null };
                publicWorkbooks.Add(selected);
            }
        }
    }

    private async Task ClaimRetainedWorkerAsync(ExcelWorkerClient worker, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await worker.ClaimAsync(cancellationToken, () =>
        {
            lock (_shutdownLock)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_shutdown is not null)
                    throw new InvalidOperationException("The service is stopping. No Excel application was claimed.");
                var pipeName = worker.Registration.PipeName;
                if (!_workers.ContainsKey(pipeName) && _workers.Count >= 32)
                    throw new InvalidOperationException("The service has reached its Excel worker limit. Close unused sessions before retrying.");
                // Once the verified pipe may receive the claim, keep it tracked even if the reply is lost.
                _workers[pipeName] = worker;
            }
        });
        ApplyWorkerSnapshot(worker);
        cancellationToken.ThrowIfCancellationRequested();
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A failed acknowledgement is diagnosed and retried. A rejected receipt cannot replace the independently verified process state.")]
    private void RetireExitedWorker(WorkerRegistration registration, bool explicitDiscovery = false)
    {
        ExcelWorkerClient worker;
        lock (_shutdownLock)
        {
            if (!_workers.TryGetValue(registration.PipeName, out var candidate)
                || candidate.Registration.Worker != registration.Worker || registration.Worker.IsAlive())
                return;
            worker = candidate;
            var confirmedExcel = explicitDiscovery ? null : registration.Excel;
            if (confirmedExcel is { } excel && !OwnedProcessGuard.TryConfirmExited(excel))
                return;
            if (!explicitDiscovery && confirmedExcel is null
                && (worker.Snapshot?.Sessions.Count > 0 || _sessions.Values.Contains(worker)
                    || _paths.Values.Any(path => path.Token == registration.PipeName)))
            {
                // The release record can arrive before the snapshot clears its session rows.
                // A released identity may still be live after an intentional detach.
                if (registration.ReleasedExcel is not { } released || !OwnedProcessGuard.TryConfirmExited(released))
                    return;
                confirmedExcel = released;
            }
            if (!explicitDiscovery)
            {
                try
                {
                    // The worker can exit after RefreshWorkerAsync checked its process.
                    // Read its last durable status again before acknowledging the record.
                    var recorded = _registry.Read(registration.PipeName);
                    // The worker could have published the application only to disk.
                    if (recorded?.Excel is { } recordedExcel && !OwnedProcessGuard.TryConfirmExited(recordedExcel))
                        return;
                    if (recorded is not null)
                    {
                        var releaseRecovered = false;
                        try
                        {
                            releaseRecovered = RecoverRecordedWorkerStatus(worker, recorded);
                        }
                        catch (InvalidOperationException ex)
                        {
                            System.Diagnostics.Trace.TraceError("Rejected exited Excel worker record: {0}", ex);
                        }
                        if (releaseRecovered)
                            ApplyWorkerSnapshot(worker);
                    }
                    _registry.Remove(registration.PipeName);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException
                    or InvalidDataException or InvalidOperationException)
                {
                    ReportWorkerRecordRemovalFailure(worker, ex);
                    return;
                }
            }
            var sessions = _sessions.Where(entry => entry.Value == worker).Select(entry => entry.Key).ToHashSet(StringComparer.Ordinal);
            if (sessions.Count > 0 && confirmedExcel is { } exited)
                RememberNativeExit(new(exited, ExcelNativeProcessState.Exited, null, null));
            var end = confirmedExcel is { } identity
                ? new EndedSession(SessionEnd.ExcelExited, identity)
                : new EndedSession(SessionEnd.WorkerExited, registration.Excel ?? registration.ReleasedExcel);
            foreach (var session in sessions)
                _closedSessions.TryAdd(session, end);
            foreach (var path in _paths.Where(entry => entry.Value.Token == registration.PipeName
                || entry.Value.SessionId is { } session && sessions.Contains(session)))
                _paths.TryRemove(path);
            foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker))
                _selections.TryRemove(selection);

            // Once the session disappears from status, requests must see its lost outcome and released claims.
            foreach (var session in sessions)
            {
                _sessions.TryRemove(session, out _);
                _unconfirmedCloseAttempts.Remove(session);
            }
            _workers.TryRemove(registration.PipeName, out _);
            _newWorkers.TryRemove(registration.PipeName, out _);
            _abandonedWorkers.TryRemove(registration.PipeName, out _);
            _publishedSnapshots.TryRemove(registration.PipeName, out _);
            AfterExitedWorkerSnapshotRemovalForTests?.Invoke(worker);
        }
        ClearUtilityWorker(worker);
    }

    private static void ReportWorkerRecordRemovalFailure(ExcelWorkerClient worker, Exception error)
    {
        worker.ReportStatusFailure("recordRemoval",
            $"The completed Excel worker record could not be removed: {error.GetType().Name}: {error.Message}", error);
        System.Diagnostics.Trace.TraceError("Excel worker {0} completion record could not be removed: {1}",
            worker.Registration.Worker.ProcessId, error);
    }

    internal Task<string?> ShutdownAsync()
    {
        lock (_shutdownLock)
            return _shutdown ??= Task.Run(StopWorkersAsync);
    }

    internal async Task<string?> RequestShutdownAsync()
    {
        var report = await ShutdownAsync();
        _host.RequestShutdown();
        _controlHost.RequestShutdown();
        return report;
    }

    internal void RevealSession(string sessionId)
    {
        if (_sessions.TryGetValue(sessionId, out var worker))
            _ = RevealSessionAsync(worker, sessionId);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Caller-abandonment reveal is diagnosed separately without replacing the original cancellation or replaying a mutation.")]
    private static async Task RevealSessionAsync(ExcelWorkerClient worker, string sessionId)
    {
        try
        {
            var response = await worker.ExecuteAsync(new ServiceRequest
            {
                Command = "window.show",
                SessionId = sessionId
            }, CancellationToken.None);
            if (!response.Success)
                System.Diagnostics.Trace.TraceError("Could not reveal retained Excel work: {0}", response.ErrorMessage);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceError("Could not reveal retained Excel work: {0}", ex);
        }
    }

    private async Task PollAsync()
    {
        while (!_stop.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(1), _stop.Token);
                await Task.WhenAll(_workers.Values.Select(PollWorkerAsync));
            }
            catch (OperationCanceledException) when (_stop.IsCancellationRequested)
            {
                return;
            }
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An unconfirmed worker process or failed acknowledgement remains for the next poll. Its status error does not stop other workers.")]
    private async Task PollWorkerAsync(ExcelWorkerClient worker)
    {
        try
        {
            var applied = await RefreshWorkerAsync(worker, _stop.Token);
            if (applied is { Registration.ReleaseConfirmed: true })
            {
                try
                {
                    RemoveReleasedWorker(worker);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    ReportWorkerRecordRemovalFailure(worker, ex);
                }
            }
            else if (!worker.Registration.Worker.IsAlive())
                RetireExitedWorker(worker.Registration);
            else if (_abandonedWorkers.ContainsKey(worker.Registration.PipeName)
                || (worker.Registration.Excel ?? worker.Registration.ReleasedExcel) is { } excel
                    && OwnedProcessGuard.TryConfirmExited(excel))
                await ReleaseEmptyWorkerAsync(worker);
        }
        catch (Win32Exception error) when (error.NativeErrorCode == WorkerProcess.ErrorAccessDenied)
        {
            worker.ReportStatusFailure("workerIdentity", $"The Excel worker process could not be verified: {error.Message}", error);
            System.Diagnostics.Trace.TraceWarning("Excel worker {0} process could not be verified: {1}",
                worker.Registration.Worker.ProcessId, error);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Shutdown reports each worker that could not safely release Excel, retaining worker/process identities for explicit handover. It never waits indefinitely or kills Excel.")]
    private async Task<string?> StopWorkersAsync()
    {
        var reports = new ConcurrentBag<string>();
        await Task.WhenAll(_workers.Values.Select(async worker =>
        {
            try
            {
                var snapshot = await worker.ReadSnapshotAsync();
                if (HasPendingWork(snapshot))
                    throw new InvalidOperationException("An Excel operation is still pending. Its worker was retained for explicit reconnection.");
                var response = await worker.ReleaseAsync(CancellationToken.None);
                var applied = ApplyWorkerSnapshot(worker);
                if (!response.Success)
                    throw new InvalidOperationException(response.ErrorMessage);
                if (response.PartialState is not null)
                    reports.Add(response.PartialState);
                else if (applied is { Registration.ReleaseConfirmed: true })
                    RemoveReleasedWorker(worker);
            }
            catch (Exception ex)
            {
                reports.Add($"Worker {worker.Registration.Worker.ProcessId}, Excel {worker.Registration.Excel?.ProcessId}: {ex.Message} " +
                    "Nothing was saved, discarded, or terminated to end this wait.");
            }
        }));
        return reports.IsEmpty ? null : string.Join(Environment.NewLine, reports);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        ShutdownAsync().GetAwaiter().GetResult();
        _stop.Cancel();
        _poll.GetAwaiter().GetResult();
        _host.Dispose();
        _controlHost.Dispose();
        _stop.Dispose();
    }
}
