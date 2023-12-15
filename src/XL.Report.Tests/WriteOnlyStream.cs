namespace XL.Report.Tests;

public sealed class WriteOnlyStream : Stream
{
    private readonly Stream stream;

    public WriteOnlyStream(Stream stream)
    {
        this.stream = stream;
    }

    public override void Close()
    {
        stream.Close();
        stream.Close();
    }

    public override async ValueTask DisposeAsync()
    {
        await stream.DisposeAsync().ConfigureAwait(false);
        await stream.DisposeAsync().ConfigureAwait(false);
    }

    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    public override int Read(Span<byte> buffer)
    {
        throw new NotSupportedException();
    }

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        throw new NotSupportedException();
    }

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = new())
    {
        throw new NotSupportedException();
    }

    public override int ReadByte()
    {
        throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        throw new NotSupportedException();
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        stream.Write(buffer);
    }

    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        await stream.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
    }

    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = new())
    {
        await stream.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
    }

    public override void WriteByte(byte value)
    {
        stream.WriteByte(value);
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        stream.Write(buffer, offset, count);
    }

    public override void Flush()
    {
        stream.Flush();
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        throw new NotSupportedException();
    }

    public override void SetLength(long value)
    {
        throw new NotSupportedException();
    }

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => stream.CanWrite;

    public override long Length => stream.Length;

    public override long Position
    {
        get => stream.Position;
        set => throw new NotSupportedException();
    }
}