using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;

namespace PowerPointTool._internal;

static class ObjectExt
{
    public static IEnumerable<object> AsEnumerable(this object obj) => obj.AsEnumerable<object>();

    public static IEnumerable<T> AsEnumerable<T>(this object obj)
    {
        if (obj == null || obj.GetType() == typeof(string))
            return null;

        return (obj as IEnumerable<T>)
            ?? (obj as System.Collections.IEnumerable)?.CastSafe<T>();
    }

    public static object GetValue(this object obj, string path)
    {
        if (obj == null || string.IsNullOrWhiteSpace(path))
            return obj;

        try
        {
            var param = Expression.Parameter(obj.GetType(), string.Empty);

            var prop = path.Trim().Split('.').Aggregate<string, Expression>(param, (r, x) =>
            {
                var startIndx = x.IndexOf('[');
                if (startIndx > 0)
                {
                    var member = Expression.PropertyOrField(r, x.Substring(0, startIndx));
                    var k = x.Substring(startIndx + 1, x.Length - startIndx - 2);

                    if (member.Type.IsArray)
                    {
                        return Expression.ArrayIndex(member, Expression.Constant(int.Parse(k)));
                    }
                    else
                    {
                        var getItemMethod = member.Type.GetProperty("Item").GetMethod;
                        var kType = getItemMethod.GetParameters()[0].ParameterType;
                        var kValue = TypeDescriptor.GetConverter(kType).ConvertFromInvariantString(k.Trim('"', '\''));

                        return Expression.Call(member, member.Type.GetProperty("Item").GetMethod,
                            Expression.Constant(kValue));
                    }
                }

                return Expression.PropertyOrField(r, x);
            });

            var getter = Expression.Lambda(prop, param);
            return getter.Compile().DynamicInvoke(obj);
        }
        catch
        {
            return null;
        }
    }
}
