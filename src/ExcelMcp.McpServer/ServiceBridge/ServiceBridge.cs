using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.McpServer.ServiceBridge;

internal interface IServiceBridgeBackend : IDisposable
{
    Task<ServiceResponse> ProcessAsync(ServiceRequest request);
    void RevealSession(string sessionId);

    /// <summary>Stops accepting work and returns a report of any Excel processes left as they are.</summary>
    string? RequestShutdown() => null;
}

internal sealed class ExcelMcpServiceBackend(Service.ExcelMcpService service) : IServiceBridgeBackend
{
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request) => service.ProcessAsync(request);
    public void RevealSession(string sessionId) => service.SessionManager.RevealForRecovery(sessionId, waitExpired: true);
    public string? RequestShutdown() => service.RequestShutdown();
    public void Dispose() => service.Dispose();
}

/// <summary>Host-owned access to the in-process Excel service.</summary>
public sealed class ServiceBridge : IDisposable
{
    private readonly SemaphoreSlim _initLock = new(1, 1);
    private readonly object _stateLock = new();
    private readonly object _disposeLock = new();
    private readonly Func<IServiceBridgeBackend> _serviceFactory;
    private readonly ILogger<ServiceBridge> _logger;
    private IServiceBridgeBackend? _current;
    private bool _disposed;
    private readonly SemaphoreSlim _dispatchSlots;

    internal ServiceBridge(Func<IServiceBridgeBackend> serviceFactory, ILogger<ServiceBridge>? logger = null,
        int maxOutstandingDispatches = 8)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(maxOutstandingDispatches, 1);
        _dispatchSlots = new SemaphoreSlim(maxOutstandingDispatches, maxOutstandingDispatches);
        _serviceFactory = serviceFactory;
        _logger = logger ?? NullLogger<ServiceBridge>.Instance;
    }

    private async Task<IServiceBridgeBackend> AcquireServiceAsync(CancellationToken cancellationToken)
    {
        lock (_stateLock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_current is not null)
                return _current;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            lock (_stateLock)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_current is not null)
                    return _current;
            }

            var backend = _serviceFactory();
            lock (_stateLock)
            {
                _current = backend;
                if (!_disposed)
                {
                    return _current;
                }
            }

            Dispose();
            throw new ObjectDisposedException(nameof(ServiceBridge));
        }
        finally
        {
            _initLock.Release();
        }
    }

    public async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (timeoutSeconds.HasValue)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(timeoutSeconds.Value, nameof(timeoutSeconds));
            ArgumentOutOfRangeException.ThrowIfGreaterThan(timeoutSeconds.Value, 4_294_967, nameof(timeoutSeconds));
        }
        lock (_stateLock)
            ObjectDisposedException.ThrowIf(_disposed, this);
        if (!_dispatchSlots.Wait(0))
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "ServiceBusy",
                ErrorMessage = "The service has reached its outstanding operation limit. Wait for active operations to finish."
            };
        var lease = new DispatchLease(_dispatchSlots);
        try
        {
            return await SendAdmittedAsync(command, sessionId, args, timeoutSeconds, cancellationToken, lease);
        }
        finally
        {
            lease.Release();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Service startup failures are logged and transported as Success=false ServiceStartup responses; caller cancellation and disposal are explicitly excluded from this boundary.")]
    private async Task<ServiceResponse> SendAdmittedAsync(string command, string? sessionId,
        object? args, int? timeoutSeconds, CancellationToken cancellationToken, DispatchLease lease)
    {
        IServiceBridgeBackend instance;
        try
        {
            instance = await AcquireServiceAsync(cancellationToken);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not ObjectDisposedException)
        {
            _logger.LogError(ex, "Failed to start the in-process Excel service.");
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "ServiceStartup",
                ErrorMessage = "Failed to start ExcelMCP Service in-process. See server logs for details.",
                ExceptionType = ex.GetType().Name
            };
        }

        cancellationToken.ThrowIfCancellationRequested();
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (timeoutSeconds.HasValue && command is not ("session.open" or "session.create" or "session.test"))
        {
            if (timeoutSeconds == 0) deadline.Cancel();
            else deadline.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds.Value));
        }

        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args is null ? null : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            CallerCancellation = deadline.Token
        };

        // Caller, dispatch and late recovery jointly retain admission until they finish.
        lease.Retain();
        var processTask = DispatchAsync(instance, request, lease, deadline.Token);

        ServiceResponse? completedCancellation = null;
        try
        {
            var response = await processTask.WaitAsync(deadline.Token);
            cancellationToken.ThrowIfCancellationRequested();
            if (!deadline.IsCancellationRequested || response.ErrorCategory != "Cancelled")
                return response;
            completedCancellation = response;
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            if (!cancellationToken.IsCancellationRequested && processTask.IsCompleted && !processTask.IsCanceled)
            {
                var response = await processTask;
                if (response.ErrorCategory != "Cancelled")
                    return response;
                completedCancellation = response;
            }
        }

        if (completedCancellation is null)
        {
            lease.Retain();
            _ = ObserveCancelledRequestAsync(instance, processTask, command, lease);
        }
        if (lease.HasStarted && !string.IsNullOrWhiteSpace(sessionId))
            RevealSession(instance, sessionId);
        cancellationToken.ThrowIfCancellationRequested();
        return new ServiceResponse
        {
            Success = false,
            Command = command,
            SessionId = sessionId,
            ErrorCategory = "Timeout",
            ErrorMessage = $"The wait expired after {timeoutSeconds} seconds. Excel work may still be running; timeout recovery did not close a session. Inspect session status for its outcome.",
            ExceptionType = nameof(TimeoutException),
            HResult = completedCancellation?.HResult,
            InnerError = completedCancellation is null
                ? null
                : $"{completedCancellation.ErrorMessage} {completedCancellation.InnerError}".Trim(),
            PartialState = completedCancellation?.PartialState
        };
    }

    private static async Task<ServiceResponse> DispatchAsync(IServiceBridgeBackend instance,
        ServiceRequest request, DispatchLease lease, CancellationToken token)
    {
        try
        {
            // Synchronous COM waits stay off the SDK request loop, behind bounded admission.
            return await Task.Run(() =>
            {
                lease.BeginDispatch(token);
                return instance.ProcessAsync(request);
            }, CancellationToken.None);
        }
        finally
        {
            lease.Release();
        }
    }

    private sealed class DispatchLease(SemaphoreSlim slots)
    {
        private int _references = 1;
        private readonly object _startLock = new();
        private bool _started;
        internal bool HasStarted
        {
            get
            {
                lock (_startLock)
                    return _started;
            }
        }
        internal void BeginDispatch(CancellationToken token)
        {
            lock (_startLock)
            {
                token.ThrowIfCancellationRequested();
                _started = true;
            }
        }
        internal void Retain() => Interlocked.Increment(ref _references);
        internal void Release()
        {
            if (Interlocked.Decrement(ref _references) == 0) slots.Release();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Observes late completion after the caller stopped waiting; errors are logged and a late opened session is revealed, never closed.")]
    private async Task ObserveCancelledRequestAsync(
        IServiceBridgeBackend instance, Task<ServiceResponse> task, string command, DispatchLease lease)
    {
        try
        {
            var response = await task;
            if (command is "session.open" or "session.create" && response.Success)
            {
                using var result = JsonDocument.Parse(response.Result
                    ?? throw new InvalidOperationException("Session creation returned no result."));
                var sessionId = result.RootElement.GetProperty("sessionId").GetString();
                if (string.IsNullOrWhiteSpace(sessionId))
                    throw new InvalidOperationException("Session creation returned no session ID.");
                RevealSession(instance, sessionId);
            }
            else if (!response.Success)
            {
                if (response.SessionId is { } pendingSession)
                    RevealSession(instance, pendingSession);
                _logger.LogWarning("Cancelled {Command} finished with error category {ErrorCategory}.",
                    command, response.ErrorCategory);
            }
        }
        catch (OperationCanceledException) when (task.IsCanceled)
        {
            _logger.LogDebug("Cancelled {Command} did not dispatch; no workbook needs recovery.", command);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Late observation of cancelled {Command} failed; workbook ownership is retained.", command);
        }
        finally
        {
            lease.Release();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Failure to reveal Excel is diagnosed without replacing the caller's cancellation or disposing the Service that owns pending work.")]
    private void RevealSession(IServiceBridgeBackend instance, string sessionId)
    {
        try
        {
            instance.RevealSession(sessionId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not reveal session {SessionId}; workbook ownership is retained.", sessionId);
        }
    }

    public void Dispose()
    {
        lock (_disposeLock)
        {
            IServiceBridgeBackend? instance;
            lock (_stateLock)
            {
                _disposed = true;
                instance = _current;
            }
            if (instance?.RequestShutdown() is { } report)
                _logger.LogWarning("{ShutdownReport}", report);
            instance?.Dispose();
            lock (_stateLock)
            {
                if (ReferenceEquals(_current, instance))
                    _current = null;
            }
        }
    }
}
