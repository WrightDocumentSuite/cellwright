using System.Runtime.InteropServices;

namespace Cellwright.Styles;

/// <summary>
/// Interns <see cref="DifferentialStyle"/> records into <c>dxf</c> indexes. Numbering starts
/// after the formats a loaded file already had, because tables and rules the model never
/// touched keep pointing at those by position.
/// </summary>
internal sealed class DifferentialStylePool(int preservedCount)
{
    private readonly Dictionary<DifferentialStyle, int> _index = [];
    private readonly List<DifferentialStyle> _styles = [];

    /// <summary>Formats added by the model, in index order after the preserved ones.</summary>
    public IReadOnlyList<DifferentialStyle> Styles => _styles;

    /// <summary>Returns the dxf index of the format, adding it on first use.</summary>
    public int Intern(DifferentialStyle style)
    {
        ref int slot = ref CollectionsMarshal.GetValueRefOrAddDefault(_index, style, out bool exists);
        if (!exists)
        {
            slot = preservedCount + _styles.Count;
            _styles.Add(style);
        }

        return slot;
    }
}
