namespace PowerPointTool.PipeTransforms;

public class EmptyPipeTransform() : BaseTransform<string>("empty")
{
    public static string DefaultValue = "EMPTY";

    protected override string ParseArgs(string[] args)
    {
        return args.Length > 0 ? args[0] : DefaultValue;
    }

    protected override object TransformItem(object obj, string emptyValue)
    {
        var str = obj?.ToString();
        return string.IsNullOrWhiteSpace(str) ? emptyValue : str;
    }
}
