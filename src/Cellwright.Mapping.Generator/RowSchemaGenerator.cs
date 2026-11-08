using System.Globalization;
using System.Text;
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

    private static readonly DiagnosticDescriptor MustBePartial = new(
        "CWM001",
        "Cellwright mapping",
        "Type '{0}' and its containing types must be declared partial to receive a generated row schema",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor UnsupportedRowShape = new(
        "CWM002",
        "Cellwright mapping",
        "Type '{0}' cannot receive a generated row schema: {1}",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor GenericTypeNotSupported = new(
        "CWM003",
        "Cellwright mapping",
        "Type '{0}' must not be generic",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor NoMappableColumns = new(
        "CWM004",
        "Cellwright mapping",
        "Type '{0}' has no mappable columns",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor MissingConstructorProperty = new(
        "CWM005",
        "Cellwright mapping",
        "Constructor parameter '{0}' of '{1}' has no matching property",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor IgnoredConstructorParameter = new(
        "CWM006",
        "Cellwright mapping",
        "Constructor parameter '{0}' of '{1}' cannot be ignored; remove it from the constructor instead",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor UnsupportedColumnType = new(
        "CWM007",
        "Cellwright mapping",
        "Property '{0}' of '{1}' has unsupported type '{2}'",
        "Cellwright.Mapping",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    private static readonly SymbolDisplayFormat FullyQualifiedWithEscapedKeywords = new(
        globalNamespaceStyle: SymbolDisplayGlobalNamespaceStyle.Included,
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
        genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters,
        miscellaneousOptions:
            SymbolDisplayMiscellaneousOptions.EscapeKeywordIdentifiers |
            SymbolDisplayMiscellaneousOptions.UseSpecialTypes);

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
            foreach (RowDiagnostic error in model.Diagnostics.Items)
            {
                Diagnostic diagnostic = error.Arg2 is not null
                    ? Diagnostic.Create(error.Descriptor, error.Location, error.Arg0, error.Arg1, error.Arg2)
                    : error.Arg1 is not null
                        ? Diagnostic.Create(error.Descriptor, error.Location, error.Arg0, error.Arg1)
                        : Diagnostic.Create(error.Descriptor, error.Location, error.Arg0);
                spc.ReportDiagnostic(diagnostic);
            }

            if (model.Diagnostics.Length == 0)
                spc.AddSource(model.HintName, SchemaEmitter.Emit(model));
        });
    }

    private static RowModel? Extract(GeneratorAttributeSyntaxContext context)
    {
        if (context.TargetSymbol is not INamedTypeSymbol type || context.TargetNode is not TypeDeclarationSyntax syntax)
            return null;

        var diagnostics = new List<RowDiagnostic>();
        string typeName = type.Name;
        Location typeLocation = syntax.Identifier.GetLocation();

        bool isPartial = IsPartial(type);
        var containers = new List<ContainerModel>();
        for (INamedTypeSymbol? container = type.ContainingType; container is not null; container = container.ContainingType)
        {
            containers.Insert(0, new ContainerModel(DeclarationKeyword(container), EscapeIdentifier(container.Name)));
            isPartial &= IsPartial(container);
            if (container.IsGenericType)
                diagnostics.Add(new RowDiagnostic(GenericTypeNotSupported, typeLocation, container.Name));
        }

        if (!isPartial)
            diagnostics.Add(new RowDiagnostic(MustBePartial, typeLocation, typeName));
        if (UnsupportedShapeReason(type) is { } shapeReason)
            diagnostics.Add(new RowDiagnostic(UnsupportedRowShape, typeLocation, typeName, shapeReason));
        if (type.IsGenericType)
            diagnostics.Add(new RowDiagnostic(GenericTypeNotSupported, typeLocation, typeName));

        bool hasHeader = true;
        foreach (KeyValuePair<string, TypedConstant> named in context.Attributes[0].NamedArguments)
        {
            if (named.Key == "HasHeader" && named.Value.Value is bool value)
                hasHeader = value;
        }

        IMethodSymbol? constructor = FindPrimaryConstructor(type);
        List<ColumnModel> columns = CollectColumns(type, constructor, diagnostics, typeLocation, out bool useConstructor);
        if (columns.Count == 0)
            diagnostics.Add(new RowDiagnostic(NoMappableColumns, typeLocation, typeName));

        return new RowModel(
            Namespace: type.ContainingNamespace is { IsGlobalNamespace: false } ns ? NamespaceName(ns) : null,
            TypeName: EscapeIdentifier(typeName),
            TypeKeyword: GetTypeKeyword(syntax),
            HintName: BuildHintName(type),
            UseConstructor: useConstructor,
            HasHeader: hasHeader,
            Containers: new EquatableArray<ContainerModel>([.. containers]),
            Columns: new EquatableArray<ColumnModel>([.. columns]),
            Diagnostics: new EquatableArray<RowDiagnostic>([.. diagnostics]));
    }

    private static bool IsPartial(INamedTypeSymbol type) => type.DeclaringSyntaxReferences.All(reference =>
        reference.GetSyntax() is TypeDeclarationSyntax declaration &&
        declaration.Modifiers.Any(modifier => modifier.IsKind(SyntaxKind.PartialKeyword)));

    private static string? UnsupportedShapeReason(INamedTypeSymbol type)
    {
        if (type.TypeKind is not TypeKind.Class and not TypeKind.Struct)
            return "only classes, structs, and records are supported";
        if (type.IsStatic)
            return "static types cannot implement IRowSchema<T>";
        if (type.IsAbstract)
            return "abstract types cannot be instantiated by Read";
        if (type.IsRefLikeType)
            return "ref-like types cannot implement IRowSchema<T>";
        if (type.IsFileLocal || ContainingTypes(type).Any(static c => c.IsFileLocal))
            return "file-local types cannot be extended from a generated source file";

        return null;
    }

    private static IEnumerable<INamedTypeSymbol> ContainingTypes(INamedTypeSymbol type)
    {
        for (INamedTypeSymbol? current = type.ContainingType; current is not null; current = current.ContainingType)
            yield return current;
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
        INamedTypeSymbol type,
        IMethodSymbol? constructor,
        List<RowDiagnostic> diagnostics,
        Location fallbackLocation,
        out bool useConstructor)
    {
        var ordered = new List<(ColumnModel Column, int SortKey)>();
        var seen = new HashSet<string>();
        int declarationIndex = 0;

        // Constructor parameters lead (positional records); properties fill in the rest.
        if (constructor is not null)
        {
            foreach (IParameterSymbol parameter in constructor.Parameters)
            {
                Location paramLocation = parameter.Locations.FirstOrDefault() ?? fallbackLocation;
                IPropertySymbol? property = FindProperty(type, parameter.Name);
                if (property is null)
                {
                    diagnostics.Add(new RowDiagnostic(MissingConstructorProperty, paramLocation, parameter.Name, type.Name));
                    continue;
                }

                if (HasAttribute(parameter, IgnoreAttribute) || HasAttribute(property, IgnoreAttribute))
                {
                    diagnostics.Add(new RowDiagnostic(IgnoredConstructorParameter, paramLocation, parameter.Name, type.Name));
                    continue;
                }

                seen.Add(property.Name);
                AddColumn(
                    ordered,
                    property.Name,
                    parameter.Type,
                    GetColumnAttribute(parameter) ?? GetColumnAttribute(property),
                    diagnostics,
                    type.Name,
                    paramLocation,
                    ref declarationIndex);
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

            Location propLocation = property.Locations.FirstOrDefault() ?? fallbackLocation;
            AddColumn(
                ordered,
                property.Name,
                property.Type,
                GetColumnAttribute(property),
                diagnostics,
                type.Name,
                propLocation,
                ref declarationIndex);
        }

        ordered.Sort(static (left, right) => left.SortKey.CompareTo(right.SortKey));
        var result = new List<ColumnModel>();
        foreach ((ColumnModel column, _) in ordered)
            result.Add(column);
        return result;
    }

    private static void AddColumn(
        List<(ColumnModel, int)> columns,
        string name,
        ITypeSymbol propertyType,
        AttributeData? attribute,
        List<RowDiagnostic> diagnostics,
        string typeName,
        Location location,
        ref int declarationIndex)
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

        if (!TryResolveKind(propertyType, out ColumnKind kind, out bool isNullable, out string? enumTypeName))
        {
            diagnostics.Add(new RowDiagnostic(
                UnsupportedColumnType,
                location,
                name,
                typeName,
                propertyType.ToDisplayString()));
            declarationIndex++;
            return;
        }

        int sortKey = order >= 0 ? order : declarationIndex + 10_000;
        columns.Add((new ColumnModel(EscapeIdentifier(name), header ?? name, kind, isNullable, format, enumTypeName), sortKey));
        declarationIndex++;
    }

    private static bool TryResolveKind(
        ITypeSymbol type,
        out ColumnKind kind,
        out bool isNullable,
        out string? enumTypeName)
    {
        isNullable = false;
        enumTypeName = null;
        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable)
        {
            isNullable = true;
            type = nullable.TypeArguments[0];
        }
        else if (type.SpecialType == SpecialType.System_String)
        {
            isNullable = type.NullableAnnotation == NullableAnnotation.Annotated;
        }

        if (type.TypeKind == TypeKind.Enum)
        {
            kind = ColumnKind.Enum;
            enumTypeName = type.ToDisplayString(FullyQualifiedWithEscapedKeywords);
            return true;
        }

        kind = type.SpecialType switch
        {
            SpecialType.System_String => ColumnKind.Text,
            SpecialType.System_Double => ColumnKind.Double,
            SpecialType.System_Single => ColumnKind.Single,
            SpecialType.System_Int16 => ColumnKind.Int16,
            SpecialType.System_Byte => ColumnKind.Byte,
            SpecialType.System_Int32 => ColumnKind.Int32,
            SpecialType.System_Int64 => ColumnKind.Int64,
            SpecialType.System_Decimal => ColumnKind.Decimal,
            SpecialType.System_Boolean => ColumnKind.Boolean,
            SpecialType.System_DateTime => ColumnKind.DateTime,
            _ => type.ToDisplayString() switch
            {
                "System.DateTimeOffset" => ColumnKind.DateTimeOffset,
                "System.DateOnly" => ColumnKind.DateOnly,
                "System.TimeOnly" => ColumnKind.TimeOnly,
                "System.TimeSpan" => ColumnKind.TimeSpan,
                "System.Guid" => ColumnKind.Guid,
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

    private static string DeclarationKeyword(INamedTypeSymbol type) => type.TypeKind == TypeKind.Interface
        ? "interface"
        : type.IsRecord
            ? type.IsValueType ? "record struct" : "record"
            : type.IsValueType ? "struct" : "class";

    private static string GetTypeKeyword(TypeDeclarationSyntax syntax) => syntax switch
    {
        RecordDeclarationSyntax record when record.ClassOrStructKeyword.IsKind(SyntaxKind.StructKeyword) => "record struct",
        RecordDeclarationSyntax => "record",
        StructDeclarationSyntax => "struct",
        _ => "class",
    };

    private static string NamespaceName(INamespaceSymbol value)
    {
        var parts = new Stack<string>();
        for (INamespaceSymbol? current = value; current is { IsGlobalNamespace: false }; current = current.ContainingNamespace)
            parts.Push(EscapeIdentifier(current.Name));

        return string.Join(".", parts);
    }

    internal static string EscapeIdentifier(string identifier) =>
        SyntaxFacts.GetKeywordKind(identifier) != SyntaxKind.None ||
        SyntaxFacts.GetContextualKeywordKind(identifier) != SyntaxKind.None
            ? "@" + identifier
            : identifier;

    private static string BuildHintName(INamedTypeSymbol type)
    {
        var parts = new Stack<string>();
        for (INamedTypeSymbol? current = type; current is not null; current = current.ContainingType)
            parts.Push(current.MetadataName);

        string metadataName = type.ContainingNamespace.IsGlobalNamespace
            ? string.Join("+", parts)
            : type.ContainingNamespace.ToDisplayString() + "::" + string.Join("+", parts);

        var hint = new StringBuilder((metadataName.Length * 4) + 24).Append("CWM_");
        foreach (char character in metadataName)
            hint.Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));

        return hint.Append(".RowSchema.g.cs").ToString();
    }
}
