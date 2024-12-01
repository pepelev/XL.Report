namespace XL.Report.Tests;

public sealed class ReadTracingStream(Stream stream) : Stream
{
    public override void Close()
    {
        stream.Close();
    }

    public override async ValueTask DisposeAsync()
    {
        await stream.DisposeAsync().ConfigureAwait(false);
    }

    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    public override int Read(Span<byte> buffer)
    {
        var startPosition = Position;
        var read = stream.Read(buffer);
        new ReadSummary(buffer.Length, startPosition, read).Print();
        return read;
    }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        var startPosition = Position;
        var read = await stream.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        new ReadSummary(count, startPosition, read).Print();
        return read;
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = new())
    {
        var startPosition = Position;
        var read =  await stream.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
        new ReadSummary(buffer.Length, startPosition, read).Print();
        return read;
    }

    public override int ReadByte()
    {
        var startPosition = Position;
        var read = stream.ReadByte();
        new ReadSummary(1, startPosition, read >= 0 ? 1 : 0).Print();
        return read;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        var startPosition = Position;
        var read = stream.Read(buffer, offset, count);
        new ReadSummary(count, startPosition, read).Print();
        return read;
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        throw new NotSupportedException();
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        throw new NotSupportedException();
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = new())
    {
        throw new NotSupportedException();
    }

    public override void WriteByte(byte value)
    {
        throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotSupportedException();
    }

    public override void Flush()
    {
        stream.Flush();
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        return stream.Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        throw new NotSupportedException();
    }

    public override bool CanRead => true;

    public override bool CanSeek => true;

    public override bool CanWrite => false;

    public override long Length => stream.Length;

    public override long Position
    {
        get => stream.Position;
        set => stream.Position = value;
    }

    private readonly record struct ReadSummary(int WantRead, long StartPosition, int ActuallyRead)
    {
        public void Print()
        {
            var message = $"READ: {ActuallyRead} bytes (want {WantRead}) {StartPosition}-{StartPosition + ActuallyRead}";
            Console.WriteLine(message);
        }
    }
}