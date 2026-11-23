using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;

namespace Avig.Kalkyl.Generators;

/// <summary>
/// Skriver mappningen för varje [KalkylRow]-typ: kolumnlistan, ToCells och FromRow. Typen
/// får IKalkylRow&lt;T&gt; implementerat i en partial-del, så att generisk kod når mappningen
/// via en statisk abstrakt egenskap utan reflektion.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class KalkylRowGenerator : IIncrementalGenerator
{
    private const string AttributeName = "Avig.Kalkyl.KalkylRowAttribute";
    private const string ColumnAttributeName = "Avig.Kalkyl.KalkylColumnAttribute";
    private const string IgnoreAttributeName = "Avig.Kalkyl.KalkylIgnoreAttribute";

    private static readonly DiagnosticDescriptor NotPartial = new(
        "KALK001",
        "Row type must be partial",
        "'{0}' must be declared partial so that Avig.Kalkyl can add its row mapping",
        "Avig.Kalkyl",
        DiagnosticSeverity.Error,
        true);

    private static readonly DiagnosticDescriptor UnsupportedType = new(
        "KALK002",
        "Property type cannot be mapped",
        "The property '{0}' has the type '{1}', which Avig.Kalkyl cannot map. Use string, a number, bool, DateTime, DateOnly, TimeOnly, Guid, an enum or CellValue, or mark it [KalkylIgnore].",
        "Avig.Kalkyl",
        DiagnosticSeverity.Error,
        true);

    private static readonly DiagnosticDescriptor NoConstructor = new(
        "KALK003",
        "Row type needs a public parameterless constructor",
        "'{0}' needs a public constructor without parameters, because rows are read with an object initializer",
        "Avig.Kalkyl",
        DiagnosticSeverity.Error,
        true);

    private static readonly DiagnosticDescriptor NotTopLevel = new(
        "KALK004",
        "Row type must be a top-level, non-generic type",
        "'{0}' is nested or generic. Declare row types at namespace level, without type parameters.",
        "Avig.Kalkyl",
        DiagnosticSeverity.Error,
        true);

    /// <inheritdoc/>
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        IncrementalValuesProvider<Generated> types = context.SyntaxProvider.ForAttributeWithMetadataName(
            AttributeName,
            static (node, _) => node is Microsoft.CodeAnalysis.CSharp.Syntax.TypeDeclarationSyntax,
            static (attributeContext, _) => Generate((INamedTypeSymbol)attributeContext.TargetSymbol, (Microsoft.CodeAnalysis.CSharp.Syntax.TypeDeclarationSyntax)attributeContext.TargetNode));

        context.RegisterSourceOutput(types, static (output, generated) =>
        {
            foreach (Diagnostic diagnostic in generated.Diagnostics)
            {
                output.ReportDiagnostic(diagnostic);
            }

            if (generated.Source is not null)
            {
                output.AddSource(generated.HintName, SourceText.From(generated.Source, Encoding.UTF8));
            }
        });
    }

    private static Generated Generate(INamedTypeSymbol type, Microsoft.CodeAnalysis.CSharp.Syntax.TypeDeclarationSyntax syntax)
    {
        List<Diagnostic> diagnostics = [];
        Location location = syntax.Identifier.GetLocation();
        string hintName = type.ToDisplayString().Replace('<', '_').Replace('>', '_') + ".Kalkyl.g.cs";

        if (syntax.Modifiers.Any(SyntaxKind.PartialKeyword) == false)
        {
            diagnostics.Add(Diagnostic.Create(NotPartial, location, type.Name));
        }

        if (type.ContainingType is not null || type.TypeParameters.Length > 0)
        {
            diagnostics.Add(Diagnostic.Create(NotTopLevel, location, type.Name));
        }

        bool hasConstructor = type.InstanceConstructors.Any(constructor => constructor.Parameters.Length == 0 && constructor.DeclaredAccessibility == Accessibility.Public);

        if (hasConstructor == false)
        {
            diagnostics.Add(Diagnostic.Create(NoConstructor, location, type.Name));
        }

        List<Property> properties = CollectProperties(type, diagnostics, location);

        if (diagnostics.Count > 0)
        {
            return new Generated(hintName, null, diagnostics);
        }

        return new Generated(hintName, Render(type, properties), diagnostics);
    }

    private static List<Property> CollectProperties(INamedTypeSymbol type, List<Diagnostic> diagnostics, Location fallback)
    {
        List<Property> properties = [];
        HashSet<string> seen = [];
        List<INamedTypeSymbol> chain = [];

        for (INamedTypeSymbol? current = type; current is not null && current.SpecialType != SpecialType.System_Object; current = current.BaseType)
        {
            chain.Insert(0, current);
        }

        int declarationIndex = 0;

        foreach (INamedTypeSymbol current in chain)
        {
            foreach (IPropertySymbol property in current.GetMembers().OfType<IPropertySymbol>())
            {
                if (property.IsStatic == true || property.IsIndexer == true || property.DeclaredAccessibility != Accessibility.Public)
                {
                    continue;
                }

                if (property.SetMethod is null || property.SetMethod.DeclaredAccessibility != Accessibility.Public || property.GetMethod is null)
                {
                    continue;
                }

                if (HasAttribute(property, IgnoreAttributeName) == true || property.Name == "EqualityContract" || seen.Add(property.Name) == false)
                {
                    continue;
                }

                Location location = property.Locations.FirstOrDefault() ?? fallback;

                if (TryClassify(property.Type, property.NullableAnnotation, out PropertyKind kind, out bool nullable, out string enumType) == false)
                {
                    diagnostics.Add(Diagnostic.Create(UnsupportedType, location, property.Name, property.Type.ToDisplayString()));

                    continue;
                }

                (string header, int order, string? numberFormat, double width) = ReadColumnAttribute(property);
                properties.Add(new Property(property.Name, header, order, declarationIndex, numberFormat, width, kind, nullable, enumType));
                declarationIndex++;
            }
        }

        return properties.OrderBy(property => property.Order).ThenBy(property => property.DeclarationIndex).ToList();
    }

    private static bool TryClassify(ITypeSymbol type, NullableAnnotation annotation, out PropertyKind kind, out bool nullable, out string enumType)
    {
        nullable = false;
        enumType = string.Empty;
        ITypeSymbol underlying = type;

        if (type is INamedTypeSymbol named && named.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            nullable = true;
            underlying = named.TypeArguments[0];
        }

        if (underlying.TypeKind == TypeKind.Enum)
        {
            kind = PropertyKind.Enum;
            enumType = underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);

            return true;
        }

        switch (underlying.SpecialType)
        {
            case SpecialType.System_String:
                kind = PropertyKind.String;
                nullable = annotation != NullableAnnotation.NotAnnotated;

                return true;

            case SpecialType.System_Int32:
                kind = PropertyKind.Int32;

                return true;

            case SpecialType.System_Int64:
                kind = PropertyKind.Int64;

                return true;

            case SpecialType.System_Int16:
                kind = PropertyKind.Int16;

                return true;

            case SpecialType.System_Byte:
                kind = PropertyKind.Byte;

                return true;

            case SpecialType.System_Double:
                kind = PropertyKind.Double;

                return true;

            case SpecialType.System_Single:
                kind = PropertyKind.Single;

                return true;

            case SpecialType.System_Decimal:
                kind = PropertyKind.Decimal;

                return true;

            case SpecialType.System_Boolean:
                kind = PropertyKind.Boolean;

                return true;

            case SpecialType.System_DateTime:
                kind = PropertyKind.DateTime;

                return true;
        }

        string name = underlying.ToDisplayString();

        switch (name)
        {
            case "System.DateOnly":
                kind = PropertyKind.DateOnly;

                return true;

            case "System.TimeOnly":
                kind = PropertyKind.TimeOnly;

                return true;

            case "System.Guid":
                kind = PropertyKind.Guid;

                return true;

            case "Avig.Kalkyl.CellValue":
                kind = PropertyKind.CellValue;
                nullable = true;

                return true;
        }

        kind = PropertyKind.String;

        return false;
    }

    private static (string Header, int Order, string? NumberFormat, double Width) ReadColumnAttribute(IPropertySymbol property)
    {
        string header = property.Name;
        int order = int.MaxValue;
        string? numberFormat = null;
        double width = 0;
        AttributeData? attribute = property.GetAttributes().FirstOrDefault(data => data.AttributeClass?.ToDisplayString() == ColumnAttributeName);

        if (attribute is null)
        {
            return (header, order, numberFormat, width);
        }

        if (attribute.ConstructorArguments.Length == 1 && attribute.ConstructorArguments[0].Value is string name && name.Trim().Length > 0)
        {
            header = name;
        }

        foreach (KeyValuePair<string, TypedConstant> argument in attribute.NamedArguments)
        {
            switch (argument.Key)
            {
                case "Order" when argument.Value.Value is int value:
                    order = value;

                    break;

                case "NumberFormat" when argument.Value.Value is string value:
                    numberFormat = value;

                    break;

                case "Width" when argument.Value.Value is double value:
                    width = value;

                    break;
            }
        }

        return (header, order, numberFormat, width);
    }

    private static bool HasAttribute(ISymbol symbol, string attributeName)
    {
        return symbol.GetAttributes().Any(data => data.AttributeClass?.ToDisplayString() == attributeName);
    }

    private static string Render(INamedTypeSymbol type, List<Property> properties)
    {
        string typeName = type.Name;
        string fullName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
        string keyword = "class";

        if (type.IsRecord == true)
        {
            keyword = "record";
        }

        StringBuilder source = new();
        source.Append("// <auto-generated/>\n");
        source.Append("// Skriven av Avig.Kalkyl.Generators för [KalkylRow]. Ändra typen, inte den här filen.\n");
        source.Append("#nullable enable\n\n");

        if (type.ContainingNamespace.IsGlobalNamespace == false)
        {
            source.Append("namespace ").Append(type.ContainingNamespace.ToDisplayString()).Append(";\n\n");
        }

        source.Append("partial ").Append(keyword).Append(' ').Append(typeName).Append(" : global::Avig.Kalkyl.IKalkylRow<").Append(fullName).Append(">\n{\n");
        source.Append("    /// <summary>The row mapping generated by Avig.Kalkyl.</summary>\n");
        source.Append("    public static global::Avig.Kalkyl.KalkylMap<").Append(fullName).Append("> KalkylMap { get; } = new __KalkylMap();\n\n");
        source.Append("    private sealed class __KalkylMap : global::Avig.Kalkyl.KalkylMap<").Append(fullName).Append(">\n    {\n");
        source.Append("        public __KalkylMap()\n            : base(new global::Avig.Kalkyl.KalkylColumnInfo[]\n            {\n");

        foreach (Property property in properties)
        {
            source.Append("                new global::Avig.Kalkyl.KalkylColumnInfo(")
                .Append(Literal(property.Header)).Append(", ")
                .Append(Bool(property.Nullable == false)).Append(", ")
                .Append(LiteralOrNull(property.NumberFormat)).Append(", ")
                .Append(property.Width.ToString("R", CultureInfo.InvariantCulture)).Append("d),\n");
        }

        source.Append("            })\n        {\n        }\n\n");
        source.Append("        public override global::System.Collections.Generic.IReadOnlyList<global::Avig.Kalkyl.Cell> ToCells(").Append(fullName).Append(" item)\n        {\n");
        source.Append("            return new global::Avig.Kalkyl.Cell[]\n            {\n");

        foreach (Property property in properties)
        {
            source.Append("                global::Avig.Kalkyl.KalkylCells.Format(").Append(WriteExpression(property)).Append(", ").Append(LiteralOrNull(property.NumberFormat)).Append("),\n");
        }

        source.Append("            };\n        }\n\n");
        source.Append("        public override global::Avig.Result<").Append(fullName).Append("> FromRow(global::Avig.Kalkyl.MappedRow row)\n        {\n");

        for (int index = 0; index < properties.Count; index++)
        {
            source.Append("            if (").Append(ReadCall(properties[index], index)).Append(" == false)\n");
            source.Append("            {\n                return row.Failure<").Append(fullName).Append(">();\n            }\n\n");
        }

        source.Append("            return new ").Append(fullName).Append("\n            {\n");

        for (int index = 0; index < properties.Count; index++)
        {
            Property property = properties[index];
            source.Append("                ").Append(property.Name).Append(" = value").Append(index.ToString(CultureInfo.InvariantCulture));

            if (property.Kind == PropertyKind.String && property.Nullable == false)
            {
                source.Append('!');
            }

            source.Append(",\n");
        }

        source.Append("            };\n        }\n    }\n}\n");

        return source.ToString();
    }

    private static string WriteExpression(Property property)
    {
        string access = "item." + property.Name;

        switch (property.Kind)
        {
            case PropertyKind.Enum:
                return "global::Avig.Kalkyl.KalkylCells.FromEnum<" + property.EnumType + ">((" + property.EnumType + "?)" + access + ")";

            case PropertyKind.String:
            case PropertyKind.CellValue:
                return "global::Avig.Kalkyl.KalkylCells.From(" + access + ")";

            default:
                return "global::Avig.Kalkyl.KalkylCells.From((" + ClrName(property.Kind) + "?)" + access + ")";
        }
    }

    private static string ReadCall(Property property, int index)
    {
        string number = index.ToString(CultureInfo.InvariantCulture);

        if (property.Kind == PropertyKind.Enum)
        {
            if (property.Nullable == true)
            {
                return "row.TryReadNullableEnum<" + property.EnumType + ">(" + number + ", out " + property.EnumType + "? value" + number + ")";
            }

            return "row.TryReadEnum<" + property.EnumType + ">(" + number + ", out " + property.EnumType + " value" + number + ")";
        }

        string typeName = ClrName(property.Kind);

        if (property.Kind == PropertyKind.String || (property.Nullable == true && property.Kind != PropertyKind.CellValue))
        {
            typeName += "?";
        }

        return "row.TryRead(" + number + ", out " + typeName + " value" + number + ")";
    }

    private static string ClrName(PropertyKind kind)
    {
        switch (kind)
        {
            case PropertyKind.String:
                return "string";

            case PropertyKind.Int32:
                return "int";

            case PropertyKind.Int64:
                return "long";

            case PropertyKind.Int16:
                return "short";

            case PropertyKind.Byte:
                return "byte";

            case PropertyKind.Double:
                return "double";

            case PropertyKind.Single:
                return "float";

            case PropertyKind.Decimal:
                return "decimal";

            case PropertyKind.Boolean:
                return "bool";

            case PropertyKind.DateTime:
                return "global::System.DateTime";

            case PropertyKind.DateOnly:
                return "global::System.DateOnly";

            case PropertyKind.TimeOnly:
                return "global::System.TimeOnly";

            case PropertyKind.Guid:
                return "global::System.Guid";

            default:
                return "global::Avig.Kalkyl.CellValue";
        }
    }

    private static string Literal(string text)
    {
        return SymbolDisplay.FormatLiteral(text, true);
    }

    private static string LiteralOrNull(string? text)
    {
        if (text is null)
        {
            return "null";
        }

        return Literal(text);
    }

    private static string Bool(bool value)
    {
        if (value == true)
        {
            return "true";
        }

        return "false";
    }

    private enum PropertyKind
    {
        String,
        Int32,
        Int64,
        Int16,
        Byte,
        Double,
        Single,
        Decimal,
        Boolean,
        DateTime,
        DateOnly,
        TimeOnly,
        Guid,
        CellValue,
        Enum,
    }

    private sealed class Property
    {
        public Property(string name, string header, int order, int declarationIndex, string? numberFormat, double width, PropertyKind kind, bool nullable, string enumType)
        {
            Name = name;
            Header = header;
            Order = order;
            DeclarationIndex = declarationIndex;
            NumberFormat = numberFormat;
            Width = width;
            Kind = kind;
            Nullable = nullable;
            EnumType = enumType;
        }

        public string Name { get; }

        public string Header { get; }

        public int Order { get; }

        public int DeclarationIndex { get; }

        public string? NumberFormat { get; }

        public double Width { get; }

        public PropertyKind Kind { get; }

        public bool Nullable { get; }

        public string EnumType { get; }
    }

    private sealed class Generated
    {
        public Generated(string hintName, string? source, List<Diagnostic> diagnostics)
        {
            HintName = hintName;
            Source = source;
            Diagnostics = diagnostics.ToImmutableArray();
        }

        public string HintName { get; }

        public string? Source { get; }

        public ImmutableArray<Diagnostic> Diagnostics { get; }
    }
}
