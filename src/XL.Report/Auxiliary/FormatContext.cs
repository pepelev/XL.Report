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

using System.Buffers;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace XL.Report.Auxiliary;

internal ref struct FormatContext(Span<char> destination)
{
    private Span<char> destination = destination;
    private bool success = true;
    private int totalCharsWritten;

    public void Write<T>(T value, ReadOnlySpan<char> format, IFormatProvider? formatProvider)
        where T : ISpanFormattable
    {
        if (!success)
        {
            return;
        }

        if (value.TryFormat(destination, out var valueWroteChars, format, formatProvider))
        {
            destination = destination[valueWroteChars..];
            totalCharsWritten += valueWroteChars;
        }
        else
        {
            success = false;
            destination = Span<char>.Empty;
        }
    }

    public void Write<T>(T value) where T : ISpanFormattable
    {
        Write(value, ReadOnlySpan<char>.Empty, CultureInfo.InvariantCulture);
    }

    public void Write<T>(T value, Format<T> format)
    {
        if (!success)
        {
            return;
        }

        if (format(value, destination, out var valueWroteChars))
        {
            destination = destination[valueWroteChars..];
            totalCharsWritten += valueWroteChars;
        }
        else
        {
            success = false;
            destination = Span<char>.Empty;
        }
    }

    public void Write(ReadOnlySpan<char> value)
    {
        if (!success)
        {
            return;
        }

        if (value.Length <= destination.Length)
        {
            value.CopyTo(destination);
            destination = destination[value.Length..];
            totalCharsWritten += value.Length;
        }
        else
        {
            success = false;
            destination = Span<char>.Empty;
        }
    }

    public bool Finish(out int charsWritten)
    {
        charsWritten = totalCharsWritten;
        return success;
    }

    public delegate bool Format<in T>(T value, Span<char> span, out int charsWritten);

    [SkipLocalsInit]
    public static string ToString<T>(
        T value,
        ReadOnlySpan<char> format = default,
        IFormatProvider? provider = null)
        where T : ISpanFormattable
    {
        const int startLength = 256;
        Span<char> buffer = stackalloc char[startLength];
        if (value.TryFormat(buffer, out var charsWritten, format, provider))
        {
            return new string(buffer[..charsWritten]);
        }

        var pool = ArrayPool<char>.Shared;
        for (var length = startLength * 2;; length = checked(length * 2))
        {
            var array = pool.Rent(length);
            if (value.TryFormat(buffer, out charsWritten, format, provider))
            {
                var result = new string(array.AsSpan(0, charsWritten));
                pool.Return(array);
                return result;
            }

            pool.Return(array);
        }
    }
}