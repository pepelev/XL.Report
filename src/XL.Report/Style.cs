namespace XL.Report;

public abstract class Style
{
    public abstract class Collection
    {
        public abstract StyleId Register(Style style);
    }

    // todo
    public class CollectionImpl : Collection
    {
        public override StyleId Register(Style style)
        {
            return new StyleId();
        }
    }
}

// todo
public class StyleImpl : Style
{
}