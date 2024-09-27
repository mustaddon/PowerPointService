namespace PowerPointTool.PipeTransforms;

internal class NumPipeTransform() : IPipeTransform
{
    public string Name => "num";

    static readonly NumberPipeTransform _transform
        = new() { IntegerMode = true };

    public object Transform(object obj, params string[] args)
        => _transform.Transform(obj, args);
}
