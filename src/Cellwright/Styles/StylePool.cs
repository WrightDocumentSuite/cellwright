using System.Runtime.InteropServices;

namespace Cellwright.Styles;

/// <summary>
/// Interns immutable <see cref="CellStyle"/> records into stable cell-format indexes.
/// Record value equality makes deduplication free. The first three entries are fixed:
/// default, date-time and duration formats used for automatic temporal styling.
/// </summary>
internal sealed class StylePool
{
    private readonly Dictionary<CellStyle, int> _index = [];
    private readonly List<CellStyle> _styles = [];

    /// <summary>Index of the fixed style applied to date-time values without an explicit style.</summary>
    public const int DateTimeStyleIndex = 1;

    /// <summary>Index of the fixed style applied to duration values without an explicit style.</summary>
    public const int TimeSpanStyleIndex = 2;

    public StylePool()
    {
        Intern(CellStyle.Default);
        Intern(CellStyle.Default with { NumberFormat = "m/d/yy h:mm" }); // Builtin 22.
        Intern(CellStyle.Default with { NumberFormat = "[h]:mm:ss" });   // Builtin 46.
    }

    /// <summary>All interned styles; the list index is the cell-format index.</summary>
    public IReadOnlyList<CellStyle> Styles => _styles;

    /// <summary>Returns the stable index of the style, adding it on first use.</summary>
    public int Intern(CellStyle style)
    {
        ref int slot = ref CollectionsMarshal.GetValueRefOrAddDefault(_index, style, out bool exists);
        if (!exists)
        {
            slot = _styles.Count;
            _styles.Add(style);
        }

        return slot;
    }
}
