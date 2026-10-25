using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "TransportLifetime")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class TransportFailureLifetimeTests
{
    [Theory]
    [InlineData("read")]
    [InlineData("read-span")]
    [InlineData("read-array-async")]
    [InlineData("read-memory-async")]
    [InlineData("write")]
    [InlineData("write-span")]
    [InlineData("write-array-async")]
    [InlineData("write-memory-async")]
    [InlineData("flush")]
    [InlineData("flush-async")]
    public async Task TerminalFailurePreservesExceptionAndSignalsOnce(string operation)
    {
        var failure = new IOException("Controlled pipe failure");
        var calls = 0;
        using var stream = new TransportLifetimeStream(new FailingStream(failure), () => calls++);
        for (var attempt = 0; attempt < 2; attempt++)
        {
            var actual = await Assert.ThrowsAsync<IOException>(() => Invoke(stream, operation));
            Assert.Same(failure, actual);
        }
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task CancellationDoesNotEndTransport()
    {
        var calls = 0;
        using var stream = new TransportLifetimeStream(new FailingStream(new OperationCanceledException()), () => calls++);
        await Assert.ThrowsAsync<OperationCanceledException>(() => Invoke(stream, "read-memory-async"));
        await Assert.ThrowsAsync<OperationCanceledException>(() => Invoke(stream, "write-memory-async"));
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task DisposedPipeSignalsAndNormalDataRemainsUnchanged()
    {
        var memory = new CountingMemoryStream();
        var calls = 0;
        var stream = new TransportLifetimeStream(memory, () => calls++);
        await stream.WriteAsync(new byte[] { 1, 2, 3 });
        memory.Position = 0;
        var actual = new byte[3];
        Assert.Equal(3, await stream.ReadAsync(actual));
        Assert.Equal(new byte[] { 1, 2, 3 }, actual);
        Assert.Equal(0, calls);
        await stream.DisposeAsync();
        stream.Dispose();
        Assert.Equal(1, memory.DisposeCalls);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => Invoke(stream, "read-memory-async"));
        Assert.Equal(1, calls);
    }

    // Each overload is a distinct adapter contract, including the legacy array API.
#pragma warning disable CA1835
    private static async Task Invoke(Stream stream, string operation)
    {
        var bytes = new byte[1];
        switch (operation)
        {
            case "read": Assert.InRange(stream.Read(bytes, 0, 1), 0, 1); break;
            case "read-span": Assert.InRange(stream.Read(bytes.AsSpan()), 0, 1); break;
            case "read-array-async": Assert.InRange(await stream.ReadAsync(bytes, 0, 1), 0, 1); break;
            case "read-memory-async": Assert.InRange(await stream.ReadAsync(bytes.AsMemory()), 0, 1); break;
            case "write": stream.Write(bytes, 0, 1); break;
            case "write-span": stream.Write(bytes.AsSpan()); break;
            case "write-array-async": await stream.WriteAsync(bytes, 0, 1); break;
            case "write-memory-async": await stream.WriteAsync(bytes.AsMemory()); break;
            case "flush": stream.Flush(); break;
            case "flush-async": await stream.FlushAsync(); break;
            default: throw new ArgumentException("Unknown test operation.", nameof(operation));
        }
    }
#pragma warning restore CA1835

    private sealed class CountingMemoryStream : MemoryStream
    {
        public int DisposeCalls { get; private set; }
        protected override void Dispose(bool disposing) { DisposeCalls++; base.Dispose(disposing); }
    }

    private sealed class FailingStream(Exception failure) : Stream
    {
        public override bool CanRead => true;
        public override bool CanWrite => true;
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw failure;
        public override int Read(Span<byte> buffer) => throw failure;
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) => ValueTask.FromException<int>(failure);
        public override void Write(byte[] buffer, int offset, int count) => throw failure;
        public override void Write(ReadOnlySpan<byte> buffer) => throw failure;
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) => ValueTask.FromException(failure);
        public override void Flush() => throw failure;
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.FromException(failure);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
