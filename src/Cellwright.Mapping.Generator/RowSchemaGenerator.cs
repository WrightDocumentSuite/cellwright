using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Cellwright.Mapping.Generator;

/// <summary>
/// Incremental generator: for every partial type marked with <c>[ExcelRow]</c> it implements
/// <c>IRowSchema&lt;T&gt;</c> — headers, span-based Read/Write and per-column styles — with
/// zero reflection at runtime.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class RowSchemaGenerator : IIncrementalGenerator
{
    private const string RowAttribute = "Cellwright.Mapping.ExcelRowAttribute";
    private const string ColumnAttribute = "Cellwright.Mapping.ExcelColumnAttribute";
    private const string IgnoreAttribute = "Cellwright.Mapping.ExcelIgnoreAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        IncrementalValuesProvider<RowModel> models = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                RowAttribute,
                static (node, _) => node is TypeDeclarationSyntax,
                static (ctx, _) => Extract(ctx))
            .Where(static model => model is not null)!;

        context.RegisterSourceOutput(models, static (spc, model) =>
        {
            foreach (string error in model.Errors.Items)
            {
                string[] parts = error.Split('|');
                spc.ReportDiagnostic(Diagnostic.Create(
                    new DiagnosticDescriptor(parts[0], "Cellwright mapping", parts[1], "Cellwright.Mapping",
                        DiagnosticSeverity.Error, isEnabledByDefault: true),
                    Location.None));
            }

            if (model.Errors.Length == 0)
                spc.AddSource($"{model.Namespace}.{model.TypeName}.RowSchema.g.cs", SchemaEmitter.Emit(model));
        });
    }

    private static RowModel? Extract(GeneratorAttributeSyntaxContext context)
    {
        if (context.TargetSymbol is not INamedTypeSymbol type || context.TargetNode is not TypeDeclarationSyntax syntax)
            return null;

        var errors = new List<string>();
        string typeName = type.Name;

        if (!syntax.Modifiers.Any(SyntaxKind.PartialKeyword))
            errors.Add($"CWM001|Type '{typeName}' must be declared partial to receive a generated row schema.");
        if (type.ContainingType is not null)
            errors.Add($"CWM002|Type '{typeName}' must not be nested inside another type.");
        if (type.IsGenericType)
            errors.Add($"CWM003|Type '{typeName}' must not be generic.");

        bool hasHeader = true;
        foreach (KeyValuePair<string, TypedConstant> named in context.Attributes[0].NamedArguments)
        {
            if (named.Key == "HasHeader" && named.Value.Value is bool value)
                hasHeader = value;
        }

        IMethodSymbol? constructor = FindPrimaryConstructor(type);
        List<ColumnModel> columns = CollectColumns(type, constructor, errors, out bool useConstructor);
        if (columns.Count == 0)
            errors.Add($"CWM004|Type '{typeName}' has no mappable columns.");

        return new RowModel(
            Namespace: type.ContainingNamespace is { IsGlobalNamespace: false } ns ? ns.ToDisplayString() : null,
            TypeName: typeName,
            TypeKeyword: GetTypeKeyword(syntax),
            UseConstructor: useConstructor,
            HasHeader: hasHeader,
            Columns: new EquatableArray<ColumnModel>([.. columns]),
            Errors: new EquatableArray<string>([.. errors]));
    }

    /// <summary>The single public constructor with parameters, ignoring the record copy constructor.</summary>
    private static IMethodSymbol? FindPrimaryConstructor(INamedTypeSymbol type)
    {
        IMethodSymbol? found = null;
        foreach (IMethodSymbol ctor in type.InstanceConstructors)
        {
            if (ctor.DeclaredAccessibility != Accessibility.Public || ctor.Parameters.Length == 0)
                continue;
            if (ctor.Parameters.Length == 1 && SymbolEqualityComparer.Default.Equals(ctor.Parameters[0].Type, type))
                continue; // Record copy constructor.
            if (found is not null)
                return null; // Ambiguous: fall back to property setters.
            found = ctor;
        }

        return found;
    }

    private static List<ColumnModel> CollectColumns(
        INamedTypeSymbol type, IMethodSymbol? constructor, List<string> errors, out bool useConstructor)
    {
        var ordered = new List<(ColumnModel Column, int SortKey)>();
        var seen = new HashSet<string>();
        int declarationIndex = 0;

        // Constructor parameters lead (positional records); properties fill in the rest.
        if (constructor is not null)
        {
            foreach (IParameterSymbol parameter in constructor.Parameters)
            {
                IPropertySymbol? property = FindProperty(type, parameter.Name);
                if (property is null)
                {
                    errors.Add($"CWM005|Constructor parameter '{parameter.Name}' of '{type.Name}' has no matching property.");
                    continue;
                }

                if (HasAttribute(parameter, IgnoreAttribute) || HasAttribute(property, IgnoreAttribute))
                {
                    errors.Add($"CWM006|Constructor parameter '{parameter.Name}' of '{type.Name}' cannot be ignored; remove it from the constructor instead.");
                    continue;
                }

                seen.Add(property.Name);
                AddColumn(ordered, property.Name, parameter.Type, GetColumnAttribute(parameter) ?? GetColumnAttribute(property),
                    errors, type.Name, ref declarationIndex);
            }
        }

        useConstructor = constructor is not null;

        foreach (ISymbol member in type.GetMembers())
        {
            if (member is not IPropertySymbol property || property.IsStatic || property.IsImplicitlyDeclared)
                continue;
            if (property.Name == "EqualityContract" || property.DeclaredAccessibility != Accessibility.Public)
                continue;
            if (property.GetMethod is null || seen.Contains(property.Name))
                continue;
            if (HasAttribute(property, IgnoreAttribute))
                continue;

            if (property.SetMethod is null)
                continue; // Read-only computed property: not mappable, silently skipped.

            AddColumn(ordered, property.Name, property.Type, GetColumnAttribute(property), errors, type.Name, ref declarationIndex);
        }

        ordered.Sort(static (left, right) => left.SortKey.CompareTo(right.SortKey));
        var result = new List<ColumnModel>();
        foreach ((ColumnModel column, _) in ordered)
            result.Add(column);
        return result;
    }

    private static void AddColumn(
        List<(ColumnModel, int)> columns, string name, ITypeSymbol propertyType, AttributeData? attribute,
        List<string> errors, string typeName, ref int declarationIndex)
    {
        string? header = null;
        string? format = null;
        int order = -1;
        if (attribute is not null)
        {
            if (attribute.ConstructorArguments.Length == 1)
                header = attribute.ConstructorArguments[0].Value as string;
            foreach (KeyValuePair<string, TypedConstant> named in attribute.NamedArguments)
            {
                if (named.Key == "Order" && named.Value.Value is int explicitOrder)
                    order = explicitOrder;
                else if (named.Key == "Format")
                    format = named.Value.Value as string;
            }
        }

        if (!TryResolveKind(propertyType, out ColumnKind kind, out bool isNullable))
        {
            errors.Add($"CWM007|Property '{name}' of '{typeName}' has unsupported type '{propertyType.ToDisplayString()}'.");
            declarationIndex++;
            return;
        }

        int sortKey = order >= 0 ? order : declarationIndex + 10_000;
        columns.Add((new ColumnModel(name, header ?? name, kind, isNullable, format), sortKey));
        declarationIndex++;
    }

    private static bool TryResolveKind(ITypeSymbol type, out ColumnKind kind, out bool isNullable)
    {
        isNullable = false;
        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable)
        {
            isNullable = true;
            type = nullable.TypeArguments[0];
        }
        else if (type.SpecialType == SpecialType.System_String)
        {
            isNullable = type.NullableAnnotation == NullableAnnotation.Annotated;
        }

        kind = type.SpecialType switch
        {
            SpecialType.System_String => ColumnKind.Text,
            SpecialType.System_Double => ColumnKind.Double,
            SpecialType.System_Int32 => ColumnKind.Int32,
            SpecialType.System_Int64 => ColumnKind.Int64,
            SpecialType.System_Decimal => ColumnKind.Decimal,
            SpecialType.System_Boolean => ColumnKind.Boolean,
            SpecialType.System_DateTime => ColumnKind.DateTime,
            _ => type.ToDisplayString() switch
            {
                "System.DateOnly" => ColumnKind.DateOnly,
                "System.TimeSpan" => ColumnKind.TimeSpan,
                "Cellwright.CellValue" => ColumnKind.CellValue,
                _ => (ColumnKind)(-1),
            },
        };

        return kind >= ColumnKind.Text;
    }

    private static IPropertySymbol? FindProperty(INamedTypeSymbol type, string parameterName)
    {
        foreach (ISymbol member in type.GetMembers())
        {
            if (member is IPropertySymbol property &&
                string.Equals(property.Name, parameterName, StringComparison.OrdinalIgnoreCase))
            {
                return property;
            }
        }

        return null;
    }

    private static bool HasAttribute(ISymbol symbol, string fullName)
    {
        foreach (AttributeData attribute in symbol.GetAttributes())
        {
            if (attribute.AttributeClass?.ToDisplayString() == fullName)
                return true;
        }

        return false;
    }

    private static AttributeData? GetColumnAttribute(ISymbol symbol)
    {
        foreach (AttributeData attribute in symbol.GetAttributes())
        {
            if (attribute.AttributeClass?.ToDisplayString() == ColumnAttribute)
                return attribute;
        }

        return null;
    }

    private static string GetTypeKeyword(TypeDeclarationSyntax syntax) => syntax switch
    {
        RecordDeclarationSyntax record when record.ClassOrStructKeyword.IsKind(SyntaxKind.StructKeyword) => "record struct",
        RecordDeclarationSyntax => "record",
        StructDeclarationSyntax => "struct",
        _ => "class",
    };
}
