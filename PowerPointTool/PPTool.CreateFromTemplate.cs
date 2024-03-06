using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace PowerPointTool;

public partial class PPTool
{
    public virtual async Task CreateFromTemplateAsync(Stream targetOutput, Stream sourceTemplate, Func<ISlideContext, object> slideModelFactory, CancellationToken cancellationToken = default)
    {
        await sourceTemplate.CopyToAsync(targetOutput, 81920, cancellationToken);
        ApplySlideModels(targetOutput, slideModelFactory);
    }

}
