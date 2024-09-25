using PowerPointTool._internal;
using System.Linq;

namespace PowerPointTool.PipeTransforms;

public abstract class BaseTransform<TArgs>(string name) : IPipeTransform
{
    public string Name => name;

    public object Transform(object obj, params string[] args)
    {
        var parsed = ParseArgs(args);

        return obj.AsEnumerable()?.Select(x => TransformItem(x, parsed))
            ?? TransformItem(obj, parsed);
    }

    protected abstract TArgs ParseArgs(string[] args);

    protected abstract object TransformItem(object obj, TArgs args);
}


public abstract class BaseTransform(string name) : BaseTransform<string[]>(name)
{
    protected override string[] ParseArgs(string[] args) => args;
}
