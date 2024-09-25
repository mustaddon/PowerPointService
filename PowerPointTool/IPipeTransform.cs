namespace PowerPointTool;

public interface IPipeTransform
{
    string Name { get; }
    object Transform(object obj, params string[] args);
}
