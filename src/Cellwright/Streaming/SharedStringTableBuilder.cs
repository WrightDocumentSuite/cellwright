using System.Runtime.InteropServices;

namespace Cellwright.Streaming;

/// <summary>
/// Deduplicating accumulator for the shared-string table written at the end of a package.
/// </summary>
internal sealed class SharedStringTableBuilder
{
    private readonly Dictionary<SharedStringItem, int> _index = [];
    private readonly List<SharedStringItem> _items = [];

    /// <summary>Number of unique strings.</summary>
    public int UniqueCount => _items.Count;

    /// <summary>Total number of cell references to the table.</summary>
    public long ReferenceCount { get; private set; }

    /// <summary>All unique strings in insertion order.</summary>
    public IReadOnlyList<SharedStringItem> Items => _items;

    /// <summary>Returns the table index for a string, adding it on first use.</summary>
    public int GetOrAdd(string value, string? richMarkup = null)
    {
        ReferenceCount++;
        var item = new SharedStringItem(value, richMarkup);
        ref int slot = ref CollectionsMarshal.GetValueRefOrAddDefault(_index, item, out bool exists);
        if (!exists)
        {
            slot = _items.Count;
            _items.Add(item);
        }

        return slot;
    }
}

/// <summary>
/// A shared-string item is identified by both its visible text and its rich runs. OOXML permits
/// two table entries with the same text but different formatting.
/// </summary>
internal readonly record struct SharedStringItem(string Text, string? RichMarkup);
