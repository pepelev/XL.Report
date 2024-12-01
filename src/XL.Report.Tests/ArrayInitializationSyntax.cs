namespace XL.Report.Tests;

[Explicit]
public sealed class ArrayInitializationSyntax
{
    [Test]
    public void TheOnlyByteTable()
    {
        var list = new List<string>();
        for (var start = 0; start < 8; start++)
        {
            for (var end = 0; end < 8; end++)
            {
                list.Add(Get(start, end));
            }
        }

        var join = string.Join(
            "," + Environment.NewLine,
            list
                .Select((@byte, index) => (Byte: @byte, Index: index))
                .GroupBy(pair => pair.Index / 8)
                .Select(group => string.Join(", ", group.Select(pair => pair.Byte)))
        );
        Console.WriteLine(join);

        string Get(int start, int end)
        {
            if (start > end)
                return "IllegalByte";

            var result = 0;
            for (var i = start; i <= end; i++)
            {
                result |= (1 << i);
            }

            var stringRepresentation = ((byte)result).ToString("b8");
            return $"0b{stringRepresentation[..4]}_{stringRepresentation[4..]}";
        }
    }

    [Test]
    public void Zeroes128()
    {
        var join = string.Join(
            $",{Environment.NewLine}",
            Enumerable
                .Repeat("0", 128)
                .Select((value, index) => (value, index))
                .GroupBy(pair => pair.index / 8)
                .Select(group => group.Select(pair => pair.value))
                .Select(group => string.Join(", ", group))
        );
        Console.WriteLine(join);
    }
}