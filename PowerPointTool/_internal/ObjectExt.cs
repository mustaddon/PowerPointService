using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;

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
            return _getValueUnsafe(obj, _pathParts(path.Trim()).GetEnumerator());
        }
        catch
        {
            return null;
        }
    }

    static readonly MethodInfo _getValueUnsafeMethod = ((Func<object, IEnumerator<string>, object>)_getValueUnsafe).Method;

    static object _getValueUnsafe(object obj, IEnumerator<string> path)
    {
        var param1 = Expression.Parameter(obj.GetType(), string.Empty);
        var param2 = Expression.Parameter(typeof(IEnumerator<string>), string.Empty);
        var expr = param1 as Expression;

        while (path.MoveNext())
        {
            if (path.Current[0] != '[')
            {
                expr = Expression.PropertyOrField(expr, path.Current);
                if (expr.Type == typeof(object))
                {
                    expr = Expression.Call(_getValueUnsafeMethod, expr, param2);
                    break;
                }
            }
            else
            {
                var k = path.Current.Substring(1, path.Current.Length - 2);

                if (expr.Type.IsArray)
                {
                    expr = Expression.ArrayIndex(expr, Expression.Constant(int.Parse(k)));
                }
                else
                {
                    var getItemMethod = expr.Type.GetProperty("Item").GetMethod;
                    var kType = getItemMethod.GetParameters()[0].ParameterType;
                    var kValue = TypeDescriptor.GetConverter(kType).ConvertFromInvariantString(k.Trim('"', '\''));

                    expr = Expression.Call(expr, getItemMethod, Expression.Constant(kValue));
                }
            }
        }

        var getter = Expression.Lambda(expr, param1, param2);
        return getter.Compile().DynamicInvoke(obj, path);
    }

    static IEnumerable<string> _pathParts(string path)
    {
        var last = 0;
        var open = false;
        var ii = path.Length - 1;

        for (var i = 0; i < path.Length; i++)
        {
            if (!open)
            {
                switch (path[i])
                {
                    case '.':
                        yield return path.Substring(last, i - last);
                        last = i + 1;
                        break;

                    case '[':
                        open = true;
                        yield return path.Substring(last, i - last);
                        last = i;
                        break;
                }
            }
            else if (path[i] == ']' && (i == ii || path[i + 1] == '.'))
            {
                open = false;
            }
        }

        if (last < path.Length)
            yield return path.Substring(last);
    }
}
