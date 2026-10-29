namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>Observes terminal transport failures without a second reader or a blocking poll.</summary>
internal sealed class TransportLifetimeStream(Stream inner, Action onEnded) : Stream
{
    private int _reported;
    private int _disposed;

    private void Ended()
    {
        if (Interlocked.Exchange(ref _reported, 1) == 0)
            onEnded();
    }

    private int Observe(int count, int requested)
    {
        if (requested > 0 && count == 0) Ended();
        return count;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        try { return Observe(inner.Read(buffer, offset, count), count); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override int Read(Span<byte> buffer)
    {
        try { return Observe(inner.Read(buffer), buffer.Length); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        try { return Observe(await inner.ReadAsync(buffer.AsMemory(offset, count), cancellationToken).ConfigureAwait(false), count); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        try { return Observe(await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false), buffer.Length); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        try { inner.Write(buffer, offset, count); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        try { inner.Write(buffer); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        try { await inner.WriteAsync(buffer.AsMemory(offset, count), cancellationToken).ConfigureAwait(false); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        try { await inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override void Flush()
    {
        try { inner.Flush(); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        try { await inner.FlushAsync(cancellationToken).ConfigureAwait(false); }
        catch (Exception error) when (error is IOException or ObjectDisposedException) { Ended(); throw; }
    }

    public override bool CanRead => inner.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => inner.CanWrite;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        try
        {
            if (disposing && Interlocked.Exchange(ref _disposed, 1) == 0) inner.Dispose();
        }
        finally { base.Dispose(disposing); }
    }
}
