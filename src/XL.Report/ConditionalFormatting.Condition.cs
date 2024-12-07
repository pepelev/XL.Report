namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public abstract void WriteAttributes(Xml xml);
        public abstract void WriteBody(Xml xml);

        public static Condition Duplicates { get; } = new Parameterless("duplicateValues");
        public static Condition Unique { get; } = new Parameterless("uniqueValues");
        public static Condition Between(Expression a, Expression b) => new Binary("between", a, b);
        public static Condition NotBetween(Expression a, Expression b) => new Binary("notBetween", a, b);
        public static Condition IsError { get; } = new Parameterless("containsErrors");
        public static Condition IsNotError { get; } = new Parameterless("notContainsErrors");
        public static Condition StartsWith(string text) => new UnaryText("beginsWith", text);
        public static Condition EndsWith(string text) => new UnaryText("endsWith", text);
        public static Condition Contains(string text) => new UnaryText("containsText", text);
        public static Condition NotContains(string text) => new UnaryText("notContainsText", text);
        public static Condition GreaterThan(Expression expression) => new Unary("greaterThan", expression);
        public static Condition GreaterThanOrEqual(Expression expression) => new Unary("greaterThanOrEqual", expression);
        public static Condition LessThan(Expression expression) => new Unary("lessThan", expression);
        public static Condition LessThanOrEqual(Expression expression) => new Unary("lessThanOrEqual", expression);
        public static Condition Equal(Expression expression) => new Unary("equal", expression);
        public static Condition NotEqual(Expression expression) => new Unary("notEqual", expression);
        public static Condition Formula(Expression expression) => new Unary("expression", expression);

        public static Condition Yesterday { get; } = new TimePeriod("yesterday");
        public static Condition Today { get; } = new TimePeriod("today");
        public static Condition Tomorrow { get; } = new TimePeriod("tomorrow");

        public static Condition Last7Days { get; } = new TimePeriod("last7Days");

        public static Condition LastWeek { get; } = new TimePeriod("lastWeek");
        public static Condition ThisWeek { get; } = new TimePeriod("thisWeek");
        public static Condition NextWeek { get; } = new TimePeriod("nextWeek");

        public static Condition LastMonth { get; } = new TimePeriod("lastMonth");
        public static Condition ThisMonth { get; } = new TimePeriod("thisMonth");
        public static Condition NextMonth { get; } = new TimePeriod("nextMonth");
    }
}