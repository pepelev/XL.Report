using XL.Report.Contents;
using String = XL.Report.Contents.String;

namespace XL.Report;

public static class SharedStringsExtensions
{
    public static String String(this SharedStrings sharedStrings, string @string) => new(@string, sharedStrings);

    public static SharedString.Force Force(this SharedStrings sharedStrings, string @string)
        => new(@string, sharedStrings);
}