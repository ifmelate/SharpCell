using System;
using System.Linq;
using System.Reflection;

namespace SharpCell.Docs;

/// <summary>The id the C# compiler gives a member in the XML documentation file.</summary>
internal static class DocId
{
    public static string Of(MemberInfo member) => member switch
    {
        Type type => "T:" + Name(type),
        ConstructorInfo constructor => $"M:{Name(constructor.DeclaringType!)}.#ctor{Parameters(constructor.GetParameters())}",
        MethodInfo method => $"M:{Name(method.DeclaringType!)}.{method.Name}{Parameters(method.GetParameters())}"
            + (method.Name is "op_Implicit" or "op_Explicit" ? "~" + Name(method.ReturnType) : ""),
        PropertyInfo property => $"P:{Name(property.DeclaringType!)}.{property.Name}{Parameters(property.GetIndexParameters())}",
        FieldInfo field => $"F:{Name(field.DeclaringType!)}.{field.Name}",
        EventInfo @event => $"E:{Name(@event.DeclaringType!)}.{@event.Name}",
        _ => throw new NotSupportedException($"No doc id for {member.MemberType} {member.Name}."),
    };

    private static string Parameters(ParameterInfo[] parameters) =>
        parameters.Length == 0 ? "" : "(" + string.Join(",", parameters.Select(p => Name(p.ParameterType))) + ")";

    private static string Name(Type type)
    {
        if (type.IsByRef)
            return Name(type.GetElementType()!) + "@";
        if (type.IsArray)
        {
            var element = Name(type.GetElementType()!);
            var rank = type.GetArrayRank();
            return rank == 1 ? element + "[]" : element + "[" + string.Join(",", Enumerable.Repeat("0:", rank)) + "]";
        }

        if (type.IsGenericType)
        {
            var definition = type.GetGenericTypeDefinition().FullName!;
            definition = definition[..definition.IndexOf('`', StringComparison.Ordinal)];
            return definition.Replace('+', '.') + "{" + string.Join(",", type.GetGenericArguments().Select(Name)) + "}";
        }

        return type.FullName!.Replace('+', '.');
    }
}
