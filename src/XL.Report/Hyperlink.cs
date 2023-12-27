namespace XL.Report;

public static class Hyperlink
{
    public static string Mailto(string email, string? subject = null) =>
        subject is { } value
            ? $"mailto:{email}?subject={value}"
            : $"mailto:{email}";
}