namespace XL.Report.Styles.Fills;

public abstract partial class Fill
{
    public abstract class Visitor<T>
    {
        public T Visit(Fill fill)
        {
            if (fill == null)
                throw new ArgumentNullException(nameof(fill));

            return fill.Accept(this);
        }

        public abstract T Visit(NoFill noFill);
        public abstract T Visit(SolidFill solidFill);
        public abstract T Visit(PatternFill patternFill);
    }
}