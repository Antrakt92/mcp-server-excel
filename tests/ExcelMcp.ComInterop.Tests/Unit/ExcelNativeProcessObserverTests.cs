using System.ComponentModel;
using System.Diagnostics;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelNativeProcessObserverTests
{
    private static readonly ExcelProcessIdentity Identity = new(123, 456);

    [Fact]
    public void Capture_CurrentProcessId_DerivesCreationTimeFromRetainedHandle()
    {
        using var process = Process.GetCurrentProcess();
        using var observer = ExcelNativeProcessObserver.Capture(process.Id);

        Assert.Equal(new ExcelProcessIdentity(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc()), observer.Identity);
        Assert.Equal(ExcelNativeProcessState.Alive, observer.Observe().State);
    }

    [Fact]
    public void Capture_CurrentProcess_UsesExactCreationIdentity()
    {
        using var process = Process.GetCurrentProcess();
        var expected = new ExcelProcessIdentity(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
        using var observer = ExcelNativeProcessObserver.Capture(expected);

        var observation = observer.Observe();

        Assert.Equal(expected, observer.Identity);
        Assert.Equal(expected, observation.Identity);
        Assert.Equal(ExcelNativeProcessState.Alive, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Null(observation.NativeError);
    }

    [Fact]
    public void Capture_CurrentProcessCreationMismatch_IsRejected()
    {
        using var process = Process.GetCurrentProcess();
        var wrong = new ExcelProcessIdentity(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc() - 1);

        Assert.Throws<InvalidOperationException>(() => ExcelNativeProcessObserver.Capture(wrong));
    }

    [Fact]
    public void Capture_RequestsOnlyQueryAndSynchronize_AndNeverReopens()
    {
        var api = new ProcessApi();
        using var observer = Capture(api);
        Assert.Equal(Identity, observer.Identity);
        Assert.Equal(0x00101000u, api.DesiredAccess);
        observer.Observe();
        observer.Observe();
        Assert.Equal(1, api.OpenCalls);
        Assert.Equal(1, api.CreationQueries);
    }

    [Fact]
    public void Capture_MismatchedCreation_DisposesHandle()
    {
        var api = new ProcessApi { CreationTime = Identity.StartedAtUtcFileTime + 1 };

        Assert.Throws<InvalidOperationException>(() => Capture(api));

        Assert.True(api.Handle.IsClosed);
        Assert.Equal(0, api.WaitCalls);
    }

    [Fact]
    public void Capture_CreationQueryFailure_PreservesErrorAndDisposesHandle()
    {
        var api = new ProcessApi { CreationQuerySucceeded = false, Error = 5 };

        var error = Assert.Throws<Win32Exception>(() => Capture(api));

        Assert.Equal(5, error.NativeErrorCode);
        Assert.True(api.Handle.IsClosed);
        Assert.Equal(0, api.WaitCalls);
    }

    [Fact]
    public void Capture_OpenFailure_PreservesErrorAndDisposesHandle()
    {
        var api = new ProcessApi { Handle = new SafeProcessHandle(IntPtr.Zero, ownsHandle: false), Error = 5 };

        var error = Assert.Throws<Win32Exception>(() => Capture(api));

        Assert.Equal(5, error.NativeErrorCode);
        Assert.True(api.Handle.IsClosed);
        Assert.Equal(0, api.CreationQueries);
    }

    [Theory]
    [InlineData(0u)]
    [InlineData(259u)]
    [InlineData(uint.MaxValue)]
    public void Observe_SignaledHandle_PreservesUnsignedExitAndCachesTerminalResult(uint exitCode)
    {
        var api = new ProcessApi { WaitResult = 0, ExitCode = exitCode };
        var observer = Capture(api);

        var observation = observer.Observe();
        api.WaitResult = 258;
        observer.Dispose();

        Assert.Equal(Identity, observation.Identity);
        Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
        Assert.Equal(exitCode, observation.ExitCode);
        Assert.Null(observation.NativeError);
        Assert.Equal(observation, observer.Observe());
        Assert.Equal(1, api.ExitQueries);
        Assert.Equal(1, api.WaitCalls);
        Assert.True(api.Handle.IsClosed);
    }

    [Fact]
    public void Observe_ExitQueryFailure_DoesNotTurnOutputIntoSuccessfulZero()
    {
        var api = new ProcessApi { WaitResult = 0, ExitQuerySucceeded = false, ExitCode = 0, Error = 5 };
        using var observer = Capture(api);

        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Equal(5, observation.NativeError);
    }

    [Fact]
    public void Observe_ExitQueryCanRecoverBeforeHandleIsDisposed()
    {
        var api = new ProcessApi { WaitResult = 0, ExitQuerySucceeded = false, Error = 5 };
        using var observer = Capture(api);
        Assert.Null(observer.Observe().ExitCode);
        api.WaitResult = uint.MaxValue;
        api.ExitQuerySucceeded = true;
        api.ExitCode = 42;

        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
        Assert.Equal(42u, observation.ExitCode);
        Assert.Null(observation.NativeError);
        Assert.Equal(1, api.WaitCalls);
    }

    [Fact]
    public void Dispose_ExitQueryFailure_RemainsExitedWithOriginalNativeError()
    {
        var api = new ProcessApi { WaitResult = 0, ExitQuerySucceeded = false, Error = 5 };
        var observer = Capture(api);
        observer.Dispose();

        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Equal(5, observation.NativeError);
        Assert.Equal(1, api.WaitCalls);
        Assert.Equal(1, api.ExitQueries);
    }

    [Fact]
    public void Observe_AndDispose_AreSerializedWithoutUsingClosedHandle()
    {
        var api = new ProcessApi();
        var observer = Capture(api);

        Parallel.For(0, 32, index =>
        {
            if (index % 2 == 0)
                observer.Observe();
            else
                observer.Dispose();
        });

        Assert.Equal(ExcelNativeProcessState.Indeterminate, observer.Observe().State);
        Assert.Equal(1, api.OpenCalls);
        Assert.True(api.Handle.IsClosed);
    }

    [Fact]
    public void Observe_WaitFailure_IsIndeterminateAndPreservesNativeError()
    {
        var api = new ProcessApi { WaitResult = uint.MaxValue, Error = 6 };
        using var observer = Capture(api);

        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Indeterminate, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Equal(6, observation.NativeError);
        Assert.Equal(0, api.ExitQueries);
    }

    [Fact]
    public void Observe_UnexpectedWaitResult_IsIndeterminateWithoutInventedError()
    {
        var api = new ProcessApi { WaitResult = 128, Error = 99 };
        using var observer = Capture(api);

        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Indeterminate, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Null(observation.NativeError);
        Assert.Equal(0, api.ExitQueries);
    }

    [Fact]
    public void Dispose_CapturesExitThatHappensAfterPreviousLiveObservation()
    {
        var api = new ProcessApi();
        var observer = Capture(api);
        Assert.Equal(ExcelNativeProcessState.Alive, observer.Observe().State);
        api.WaitResult = 0;
        api.ExitCode = 17;

        observer.Dispose();
        observer.Dispose();
        var observation = observer.Observe();

        Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
        Assert.Equal(17u, observation.ExitCode);
        Assert.Null(observation.NativeError);
        Assert.Equal(2, api.WaitCalls);
        Assert.Equal(1, api.ExitQueries);
        Assert.True(api.Handle.IsClosed);
    }

    [Fact]
    public void Dispose_LiveProcessCannotRemainAnAuthoritativeLiveObservation()
    {
        var api = new ProcessApi();
        var observer = Capture(api);
        observer.Dispose();
        var observation = observer.Observe();

        Assert.Equal(Identity, observation.Identity);
        Assert.Equal(ExcelNativeProcessState.Indeterminate, observation.State);
        Assert.Null(observation.ExitCode);
        Assert.Equal(6, observation.NativeError);
        Assert.True(api.Handle.IsClosed);
        Assert.Equal(1, api.WaitCalls);
    }

    private static ExcelNativeProcessObserver Capture(ProcessApi api) =>
        ExcelNativeProcessObserver.Capture(Identity.ProcessId, Identity.StartedAtUtcFileTime, api);

    private sealed class ProcessApi : ExcelNativeProcessObserver.IProcessApi
    {
        internal SafeProcessHandle Handle { get; init; } = new(new IntPtr(42), ownsHandle: false);
        internal long CreationTime { get; init; } = Identity.StartedAtUtcFileTime;
        internal bool CreationQuerySucceeded { get; init; } = true;
        internal uint WaitResult { get; set; } = 258;
        internal bool ExitQuerySucceeded { get; set; } = true;
        internal uint ExitCode { get; set; }
        internal int Error { get; init; }
        internal uint DesiredAccess { get; private set; }
        internal int OpenCalls { get; private set; }
        internal int CreationQueries { get; private set; }
        internal int WaitCalls { get; private set; }
        internal int ExitQueries { get; private set; }

        public SafeProcessHandle Open(int processId, uint desiredAccess, out int error)
        {
            Assert.Equal(Identity.ProcessId, processId);
            DesiredAccess = desiredAccess;
            OpenCalls++;
            error = Error;
            return Handle;
        }

        public bool GetCreationTime(SafeProcessHandle handle, out long creationTime, out int error)
        {
            Assert.Same(Handle, handle);
            CreationQueries++;
            creationTime = CreationTime;
            error = Error;
            return CreationQuerySucceeded;
        }

        public uint Wait(SafeProcessHandle handle, out int error)
        {
            Assert.Same(Handle, handle);
            Assert.False(handle.IsClosed);
            WaitCalls++;
            error = Error;
            return WaitResult;
        }

        public bool GetExitCode(SafeProcessHandle handle, out uint exitCode, out int error)
        {
            Assert.Same(Handle, handle);
            Assert.False(handle.IsClosed);
            ExitQueries++;
            exitCode = ExitCode;
            error = Error;
            return ExitQuerySucceeded;
        }
    }
}
