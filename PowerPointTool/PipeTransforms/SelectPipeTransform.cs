using PowerPointTool._internal;

namespace PowerPointTool.PipeTransforms;

public class SelectPipeTransform() : BaseTransform<string>("select")
{
    protected override string ParseArgs(string[] args)
    {
        return args.Length > 0 ? args[0] : null;
    }

    protected override object TransformItem(object obj, string prop)
    {
        return obj.GetValue(prop);
    }
}
