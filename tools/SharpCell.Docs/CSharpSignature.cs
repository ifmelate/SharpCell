using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;

namespace SharpCell.Docs;

/// <summary>How a type or member is declared, written as C#.</summary>
internal static class CSharpSignature
{
    private static readonly NullabilityInfoContext Nullability = new();

    private static readonly Dictionary<Type, string> Keywords = new()
    {
        [typeof(void)] = "void", [typeof(object)] = "object", [typeof(string)] = "string", [typeof(bool)] = "bool",
        [typeof(char)] = "char", [typeof(byte)] = "byte", [typeof(short)] = "short", [typeof(int)] = "int",
        [typeof(long)] = "long", [typeof(uint)] = "uint", [typeof(ulong)] = "ulong", [typeof(float)] = "float",
        [typeof(double)] = "double", [typeof(decimal)] = "decimal",
    };

    private static readonly Dictionary<string, string> Operators = new(StringComparer.Ordinal)
    {
        ["op_Equality"] = "==", ["op_Inequality"] = "!=", ["op_LessThan"] = "<", ["op_GreaterThan"] = ">",
        ["op_LessThanOrEqual"] = "<=", ["op_GreaterThanOrEqual"] = ">=", ["op_Addition"] = "+", ["op_Subtraction"] = "-",
    };

    public static string Declaration(Type type)
    {
        var sb = new StringBuilder("public ");
        if (type.IsEnum)
            return sb.Append("enum ").Append(type.Name).ToString();
        if (type.IsValueType)
        {
            if (type.IsDefined(typeof(IsReadOnlyAttribute), false))
                sb.Append("readonly ");
            return sb.Append("struct ").Append(type.Name).ToString();
        }

        if (type.IsAbstract && type.IsSealed)
            sb.Append("static ");
        else if (type.IsSealed)
            sb.Append("sealed ");
        else if (type.IsAbstract)
            sb.Append("abstract ");
        sb.Append("class ").Append(type.Name);
        if (type.BaseType is { } baseType && baseType != typeof(object))
            sb.Append(" : ").Append(TypeName(baseType));
        return sb.ToString();
    }

    public static string Of(MemberInfo member) => member switch
    {
        ConstructorInfo constructor => $"public {constructor.DeclaringType!.Name}({Parameters(constructor.GetParameters(), false)})",
        MethodInfo method => Method(method),
        PropertyInfo property => Property(property),
        FieldInfo { DeclaringType.IsEnum: true } field =>
            $"{field.Name} = {Convert.ToInt64(field.GetRawConstantValue(), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)}",
        FieldInfo field => $"public {(field.IsLiteral ? "const" : field.IsStatic ? "static readonly" : "readonly")} {TypeName(field.FieldType, Nullability.Create(field))} {field.Name}",
        _ => throw new NotSupportedException($"No signature for {member.MemberType} {member.Name}."),
    };

    /// <summary>The member as a heading: name and parameter types, enough to tell overloads apart.</summary>
    public static string Heading(MemberInfo member) => member switch
    {
        ConstructorInfo constructor => $"{constructor.DeclaringType!.Name}({ParameterTypes(constructor.GetParameters())})",
        MethodInfo { Name: "op_Implicit" or "op_Explicit" } conversion =>
            $"{(conversion.Name == "op_Implicit" ? "implicit" : "explicit")} operator {TypeName(conversion.ReturnType)}({ParameterTypes(conversion.GetParameters())})",
        MethodInfo method when Operators.TryGetValue(method.Name, out var symbol) => $"operator {symbol}({ParameterTypes(method.GetParameters())})",
        MethodInfo method => $"{method.Name}({ParameterTypes(method.GetParameters())})",
        PropertyInfo property when property.GetIndexParameters().Length > 0 => $"this[{ParameterTypes(property.GetIndexParameters())}]",
        _ => member.Name,
    };

    /// <summary>A stable HTML id for the member, unique among overloads and operators.</summary>
    public static string Anchor(MemberInfo member)
    {
        var parameters = member switch
        {
            MethodBase method => method.GetParameters(),
            PropertyInfo property => property.GetIndexParameters(),
            _ => [],
        };
        var name = member is ConstructorInfo ? "ctor" : member.Name;
        return Slug(string.Join(" ", parameters.Select(p => TypeName(p.ParameterType)).Prepend(name)));
    }

    public static string TypeName(Type type, NullabilityInfo? nullability = null)
    {
        if (type.IsByRef)
            return TypeName(type.GetElementType()!, nullability);
        if (Nullable.GetUnderlyingType(type) is { } underlying)
            return TypeName(underlying) + "?";

        string name;
        if (type.IsArray)
            name = TypeName(type.GetElementType()!, nullability?.ElementType) + "[" + new string(',', type.GetArrayRank() - 1) + "]";
        else if (Keywords.TryGetValue(type, out var keyword))
            name = keyword;
        else if (type.IsGenericType)
        {
            var arguments = type.GetGenericArguments();
            name = type.Name[..type.Name.IndexOf('`', StringComparison.Ordinal)] + "<"
                + string.Join(", ", arguments.Select((a, i) => TypeName(a, nullability?.GenericTypeArguments.ElementAtOrDefault(i)))) + ">";
        }
        else
            name = type.Name;

        return !type.IsValueType && nullability?.ReadState == NullabilityState.Nullable ? name + "?" : name;
    }

    private static string Method(MethodInfo method)
    {
        var modifiers = "public " + (method.IsStatic ? "static " : "")
            + (!method.IsStatic && method.GetBaseDefinition().DeclaringType != method.DeclaringType ? "override " : "");
        var parameters = Parameters(method.GetParameters(), method.IsDefined(typeof(ExtensionAttribute), false));
        if (method.Name is "op_Implicit" or "op_Explicit")
            return $"{modifiers}{(method.Name == "op_Implicit" ? "implicit" : "explicit")} operator {TypeName(method.ReturnType)}({parameters})";
        var returns = TypeName(method.ReturnType, Nullability.Create(method.ReturnParameter));
        return Operators.TryGetValue(method.Name, out var symbol)
            ? $"{modifiers}{returns} operator {symbol}({parameters})"
            : $"{modifiers}{returns} {method.Name}({parameters})";
    }

    private static string Property(PropertyInfo property)
    {
        var accessor = property.GetMethod ?? property.SetMethod!;
        var sb = new StringBuilder("public ");
        if (accessor.IsStatic)
            sb.Append("static ");
        sb.Append(TypeName(property.PropertyType, Nullability.Create(property))).Append(' ');
        var index = property.GetIndexParameters();
        sb.Append(index.Length > 0 ? $"this[{Parameters(index, false)}]" : property.Name);
        sb.Append(" { ");
        if (property.GetMethod is { IsPublic: true })
            sb.Append("get; ");
        if (property.SetMethod is { IsPublic: true })
            sb.Append(property.SetMethod.ReturnParameter.GetRequiredCustomModifiers().Contains(typeof(IsExternalInit)) ? "init; " : "set; ");
        return sb.Append('}').ToString();
    }

    private static string Parameters(ParameterInfo[] parameters, bool extension) =>
        string.Join(", ", parameters.Select((p, i) => (extension && i == 0 ? "this " : "") + Parameter(p)));

    private static string Parameter(ParameterInfo parameter)
    {
        var prefix = parameter.IsOut ? "out " : parameter.ParameterType.IsByRef ? "ref " : "";
        var text = $"{prefix}{TypeName(parameter.ParameterType, Nullability.Create(parameter))} {parameter.Name}";
        return parameter.HasDefaultValue ? text + " = " + DefaultValue(parameter) : text;
    }

    private static string DefaultValue(ParameterInfo parameter) => parameter.DefaultValue switch
    {
        null => parameter.ParameterType.IsValueType ? "default" : "null",
        bool b => b ? "true" : "false",
        string s => "\"" + s.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"",
        Enum e => $"{parameter.ParameterType.Name}.{e}",
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        var other => other.ToString() ?? "default",
    };

    private static string ParameterTypes(ParameterInfo[] parameters) => string.Join(", ", parameters.Select(p => TypeName(p.ParameterType)));

    private static string Slug(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text.ToLowerInvariant())
        {
            if (c is >= 'a' and <= 'z' or >= '0' and <= '9')
                sb.Append(c);
            else if (sb.Length > 0 && sb[^1] != '-')
                sb.Append('-');
        }

        return sb.ToString().Trim('-');
    }
}
