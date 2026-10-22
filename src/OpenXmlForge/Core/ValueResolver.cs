using System.Collections;
using System.Collections.Concurrent;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;
using OpenXmlForge.Models;

namespace OpenXmlForge.Core;

internal static class ValueResolver
{
    private static readonly ConcurrentDictionary<Type, PropertyAccessor[]> PropertyCache = new();

    internal readonly record struct PropertyAccessor(string Name, Func<object, object?> Getter);

    public static Dictionary<string, object?> NormalizeRoot(object? value)
    {
        if (value is null)
        {
            return new Dictionary<string, object?>(StringComparer.Ordinal);
        }

        if (value is Dictionary<string, object?> exactDict && ReferenceEquals(exactDict.Comparer, StringComparer.Ordinal))
        {
            return new Dictionary<string, object?>(exactDict, StringComparer.Ordinal);
        }

        if (value is IDictionary<string, object?> genericDict)
        {
            var copy = new Dictionary<string, object?>(genericDict.Count, StringComparer.Ordinal);
            foreach (var kvp in genericDict)
            {
                copy[kvp.Key] = kvp.Value;
            }

            return copy;
        }

        if (value is IReadOnlyDictionary<string, object?> readOnlyDict)
        {
            var copy = new Dictionary<string, object?>(readOnlyDict.Count, StringComparer.Ordinal);
            foreach (var kvp in readOnlyDict)
            {
                copy[kvp.Key] = kvp.Value;
            }

            return copy;
        }

        if (value is IDictionary nonGenericDict)
        {
            var copy = new Dictionary<string, object?>(nonGenericDict.Count, StringComparer.Ordinal);
            foreach (DictionaryEntry entry in nonGenericDict)
            {
                if (entry.Key is not null)
                {
                    copy[entry.Key.ToString()!] = entry.Value;
                }
            }

            return copy;
        }

        if (IsRowCollection(value))
        {
            throw new ArgumentException("The root template data parameter cannot be a collection type.", nameof(value));
        }

        return ObjectToDictionary(value);
    }

    public static Dictionary<string, object?> ObjectToDictionary(object? item)
    {
        if (item is null)
        {
            return new Dictionary<string, object?>(StringComparer.Ordinal);
        }

        if (item is IDictionary<string, object?> genericDict)
        {
            var copy = new Dictionary<string, object?>(genericDict.Count, StringComparer.Ordinal);
            foreach (var kvp in genericDict)
            {
                copy[kvp.Key] = kvp.Value;
            }

            return copy;
        }

        if (item is IReadOnlyDictionary<string, object?> roDict)
        {
            var copy = new Dictionary<string, object?>(roDict.Count, StringComparer.Ordinal);
            foreach (var kvp in roDict)
            {
                copy[kvp.Key] = kvp.Value;
            }

            return copy;
        }

        if (item is IDictionary nonGenericDict)
        {
            var copy = new Dictionary<string, object?>(nonGenericDict.Count, StringComparer.Ordinal);
            foreach (DictionaryEntry entry in nonGenericDict)
            {
                if (entry.Key is not null)
                {
                    copy[entry.Key.ToString()!] = entry.Value;
                }
            }

            return copy;
        }

        var accessors = GetAccessors(item.GetType());
        var result = new Dictionary<string, object?>(accessors.Length, StringComparer.Ordinal);

        foreach (var accessor in accessors)
        {
            result[accessor.Name] = accessor.Getter(item);
        }

        return result;
    }

    public static bool TryResolveValue(
        IReadOnlyDictionary<string, object?> scope,
        string rawPath,
        out object? value)
    {
        var path = rawPath.Trim();

        if (TryResolveExactOrDotted(scope, path, out value))
        {
            return true;
        }

        if (path.AsSpan().IndexOfAny(" \t\r\n") >= 0)
        {
            var compactPath = RemoveWhitespace(path);
            if (TryResolveExactOrDotted(scope, compactPath, out value))
            {
                return true;
            }
        }

        value = null;
        return false;
    }

    public static object? ResolveValueOrDefault(
        IReadOnlyDictionary<string, object?> scope,
        string rawPath)
        => TryResolveValue(scope, rawPath, out var value) ? value : null;

    public static bool IsRowCollection(object? value)
    {
        if (value is null or string or byte[] or char[] or Stream
            or IDictionary or IReadOnlyDictionary<string, object?> or IReadOnlyDictionary<string, object>
            or ExpandoObject
            or WordPicture or IEnumerable<WordPicture>
            or WordColorText or IEnumerable<WordColorText>
            or WordHyperLink or IEnumerable<WordHyperLink>
            or WordForeach or IEnumerable<WordForeach>
            or string[] or IList<string> or IReadOnlyList<string> or IEnumerable<string>)
        {
            return false;
        }

        return value is IEnumerable;
    }

    public static bool IsStringEnumerable(object? value)
        => value is not string && value is IEnumerable<string>;

    public static void CollectPictures(object? node, HashSet<WordPicture> pictures)
    {
        switch (node)
        {
            case null:
                return;

            case WordPicture pic:
                pictures.Add(pic);
                return;

            case IEnumerable<WordPicture> picList:
                foreach (var pic in picList)
                {
                    if (pic is not null)
                    {
                        pictures.Add(pic);
                    }
                }

                return;

            case IDictionary<string, object?> genericDict:
                foreach (var val in genericDict.Values)
                {
                    CollectPictures(val, pictures);
                }

                return;

            case IDictionary dict:
                foreach (DictionaryEntry entry in dict)
                {
                    CollectPictures(entry.Value, pictures);
                }

                return;

            case IEnumerable enumerable and not string and not byte[] and not char[] and not Stream:
                foreach (var item in enumerable)
                {
                    CollectPictures(item, pictures);
                }

                return;

            default:
                var type = node.GetType();
                if (!type.IsPrimitive && !type.IsEnum && type != typeof(string) && type != typeof(DateTime) &&
                    type != typeof(DateTimeOffset) && type != typeof(decimal) && type != typeof(Guid) &&
                    !typeof(Stream).IsAssignableFrom(type))
                {
                    foreach (var accessor in GetAccessors(type))
                    {
                        CollectPictures(accessor.Getter(node), pictures);
                    }
                }

                return;
        }
    }

    internal static PropertyAccessor[] GetAccessors(Type type)
        => PropertyCache.GetOrAdd(type, static t =>
        {
            var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var list = new List<PropertyAccessor>(props.Length);

            foreach (var prop in props)
            {
                if (!prop.CanRead || prop.GetIndexParameters().Length > 0)
                {
                    continue;
                }

                var param = Expression.Parameter(typeof(object), "obj");
                var castObj = Expression.Convert(param, t);
                var propAccess = Expression.Property(castObj, prop);
                var castResult = Expression.Convert(propAccess, typeof(object));
                var getter = Expression.Lambda<Func<object, object?>>(castResult, param).Compile();

                list.Add(new PropertyAccessor(prop.Name, getter));
            }

            return [.. list];
        });

    private static bool TryResolveExactOrDotted(
        IReadOnlyDictionary<string, object?> scope,
        string path,
        out object? value)
    {
        if (scope.TryGetValue(path, out value))
        {
            return true;
        }

        var dotIndex = path.IndexOf('.');
        if (dotIndex < 0)
        {
            value = null;
            return false;
        }

        // Support keys in scope that themselves contain dots (e.g. "A.B" -> { "C": val })
        var searchPos = path.Length;
        while ((searchPos = path.LastIndexOf('.', searchPos - 1)) > 0)
        {
            var prefix = path[..searchPos];
            if (scope.TryGetValue(prefix, out var prefixObj))
            {
                var remainder = path[(searchPos + 1)..];
                if (prefixObj is null)
                {
                    value = null;
                    return true;
                }

                return TryResolveFromObject(prefixObj, remainder.AsSpan(), out value);
            }
        }

        value = null;
        return false;
    }

    private static bool TryResolveFromObject(
        object? current,
        ReadOnlySpan<char> remainingPath,
        out object? value)
    {
        while (!remainingPath.IsEmpty)
        {
            if (current is null)
            {
                value = null;
                return true;
            }

            if (IsRowCollection(current))
            {
                value = null;
                return true;
            }

            var dotIdx = remainingPath.IndexOf('.');
            var segment = dotIdx < 0 ? remainingPath : remainingPath[..dotIdx];
            var segmentStr = segment.ToString();

            var target = current;
            if (!TryGetMember(target, segmentStr, out current))
            {
                // Also check if target is a dictionary that contains the entire remainingPath with dots
                if (dotIdx >= 0 && TryGetMember(target, remainingPath.ToString(), out value))
                {
                    return true;
                }

                value = null;
                return false;
            }

            if (dotIdx < 0)
            {
                break;
            }

            remainingPath = remainingPath[(dotIdx + 1)..];
        }

        value = current;
        return true;
    }

    private static bool TryGetMember(object target, string memberName, out object? value)
    {
        if (target is IDictionary<string, object?> genericDict)
        {
            return genericDict.TryGetValue(memberName, out value);
        }

        if (target is IReadOnlyDictionary<string, object?> roDict)
        {
            return roDict.TryGetValue(memberName, out value);
        }

        if (target is IDictionary nonGenericDict)
        {
            if (nonGenericDict.Contains(memberName))
            {
                value = nonGenericDict[memberName];
                return true;
            }

            value = null;
            return false;
        }

        var accessors = GetAccessors(target.GetType());
        foreach (var accessor in accessors)
        {
            if (string.Equals(accessor.Name, memberName, StringComparison.Ordinal))
            {
                value = accessor.Getter(target);
                return true;
            }
        }

        value = null;
        return false;
    }

    private static string RemoveWhitespace(string input)
    {
        Span<char> buffer = stackalloc char[input.Length];
        var pos = 0;
        foreach (var ch in input)
        {
            if (!char.IsWhiteSpace(ch))
            {
                buffer[pos++] = ch;
            }
        }

        return new string(buffer[..pos]);
    }
}
