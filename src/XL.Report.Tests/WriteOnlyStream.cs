#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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