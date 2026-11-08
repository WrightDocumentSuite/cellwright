using Microsoft.CodeAnalysis;

namespace Cellwright.Mapping.Generator;

/// <summary>Conversion category of one mapped column.</summary>
internal enum ColumnKind
{
    Text,
    Double,
    Single,
    Int16,
    Byte,
    Int32,
    Int64,
    Decimal,
    Boolean,
    DateTime,
    DateTimeOffset,
    DateOnly,
    TimeOnly,
    TimeSpan,
    Guid,
    Enum,
    CellValue,
}

/// <summary>One enclosing partial type declaration for a nested row model.</summary>
internal readonly record struct ContainerModel(string Keyword, string Name);

/// <summary>One diagnostic produced while analyzing a row model.</summary>
internal readonly record struct RowDiagnostic(
    DiagnosticDescriptor Descriptor,
    Location Location,
    string Arg0,
    string? Arg1 = null,
    string? Arg2 = null);

/// <summary>One mapped column extracted from a property or positional parameter.</summary>
internal sealed record ColumnModel(
    string PropertyName,
    string Header,
    ColumnKind Kind,
    bool IsNullable,
    string? Format,
    string? EnumTypeName);

/// <summary>Everything the emitter needs to generate one row schema.</summary>
internal sealed record RowModel(
    string? Namespace,
    string TypeName,
    string TypeKeyword,
    string HintName,
    bool UseConstructor,
    bool HasHeader,
    EquatableArray<ContainerModel> Containers,
    EquatableArray<ColumnModel> Columns,
    EquatableArray<RowDiagnostic> Diagnostics);

/// <summary>Array wrapper with sequence equality, required for incremental caching.</summary>
internal readonly struct EquatableArray<T>(T[] items) : IEquatable<EquatableArray<T>> where T : IEquatable<T>
{
    private readonly T[] _items = items;

    public T[] Items => _items ?? [];

    public int Length => Items.Length;

    public T this[int index] => Items[index];

    public bool Equals(EquatableArray<T> other)
    {
        T[] left = Items;
        T[] right = other.Items;
        if (left.Length != right.Length)
            return false;
        for (int i = 0; i < left.Length; i++)
        {
            if (!left[i].Equals(right[i]))
                return false;
        }

        return true;
    }

    public override bool Equals(object? obj) => obj is EquatableArray<T> other && Equals(other);

    public override int GetHashCode()
    {
        int hash = 17;
        foreach (T item in Items)
            hash = unchecked(hash * 31 + item.GetHashCode());
        return hash;
    }
}
