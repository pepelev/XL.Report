namespace XL.Report;

public interface IKeyed<out T>
{
    T Key { get; }
}