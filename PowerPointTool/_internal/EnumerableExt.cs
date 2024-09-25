using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace PowerPointTool._internal;

static class EnumerableExt
{
    public static IEnumerable<T> CastSafe<T>(this IEnumerable source)
    {
        try
        {
            return source.Cast<T>();
        }
        catch { }

        return null;
    }


}
