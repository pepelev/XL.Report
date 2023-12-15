namespace XL.Report;

public interface IUnit<out T>
{
    public T Write(SheetWindow window);
}