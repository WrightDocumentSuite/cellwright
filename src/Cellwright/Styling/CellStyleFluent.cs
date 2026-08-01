namespace Cellwright;

/// <summary>
/// Fluent builders for <see cref="CellStyle"/> implemented as C# 14 extension members,
/// keeping the record itself a pure data type:
/// <c>CellStyle.Default.Bold().Fill(0xFF2278D4).NumberFormat("#,##0.00")</c>.
/// </summary>
public static class CellStyleFluent
{
    /// <summary>Extension members for building styles immutably.</summary>
    extension(CellStyle style)
    {
        /// <summary>Whether the style differs from the workbook default.</summary>
        public bool IsDefault => style == CellStyle.Default;

        /// <summary>Returns the style with bold toggled.</summary>
        public CellStyle Bold(bool value = true) => style with { Font = style.Font with { Bold = value } };

        /// <summary>Returns the style with italic toggled.</summary>
        public CellStyle Italic(bool value = true) => style with { Font = style.Font with { Italic = value } };

        /// <summary>Returns the style with a single underline toggled.</summary>
        public CellStyle Underline(bool value = true) =>
            style with { Font = style.Font with { Underline = value ? UnderlineStyle.Single : UnderlineStyle.None } };

        /// <summary>Returns the style with a particular kind of underline.</summary>
        public CellStyle Underline(UnderlineStyle underline) =>
            style with { Font = style.Font with { Underline = underline } };

        /// <summary>Returns the style with strikethrough toggled.</summary>
        public CellStyle Strikethrough(bool value = true) => style with { Font = style.Font with { Strikethrough = value } };

        /// <summary>Returns the style with another font family.</summary>
        public CellStyle FontName(string name)
        {
            ArgumentException.ThrowIfNullOrEmpty(name);
            return style with { Font = style.Font with { Name = name } };
        }

        /// <summary>Returns the style with another font size in points.</summary>
        public CellStyle FontSize(double points)
        {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(points);
            return style with { Font = style.Font with { Size = points } };
        }

        /// <summary>Returns the style with another text color.</summary>
        public CellStyle FontColor(ExcelColor color) => style with { Font = style.Font with { Color = color } };

        /// <summary>Returns the style with a solid fill.</summary>
        public CellStyle Fill(ExcelColor color) => style with { Fill = FillStyle.Solid(color) };

        /// <summary>Returns the style with a solid fill from an ARGB value, e.g. <c>0xFF2278D4</c>.</summary>
        public CellStyle Fill(uint argb) => style with { Fill = FillStyle.Solid(ExcelColor.FromArgb(argb)) };

        /// <summary>Returns the style with the same border on all four edges.</summary>
        public CellStyle Border(BorderEdge edge) => style with { Border = BorderStyle.All(edge) };

        /// <summary>Returns the style with a thin border on all four edges.</summary>
        public CellStyle Border() => style with { Border = BorderStyle.All(BorderEdge.Thin) };

        /// <summary>Returns the style with another number format code.</summary>
        public CellStyle NumberFormat(string code)
        {
            ArgumentNullException.ThrowIfNull(code);
            return style with { NumberFormat = code };
        }

        /// <summary>Returns the style with horizontal alignment.</summary>
        public CellStyle Align(HorizontalAlignment horizontal) =>
            style with { Alignment = style.Alignment with { Horizontal = horizontal } };

        /// <summary>Returns the style with vertical alignment.</summary>
        public CellStyle Align(VerticalAlignment vertical) =>
            style with { Alignment = style.Alignment with { Vertical = vertical } };

        /// <summary>Returns the style centered both ways.</summary>
        public CellStyle Center() => style with
        {
            Alignment = style.Alignment with { Horizontal = HorizontalAlignment.Center, Vertical = VerticalAlignment.Center },
        };

        /// <summary>Returns the style with text wrapping.</summary>
        public CellStyle Wrap(bool value = true) => style with { Alignment = style.Alignment with { WrapText = value } };

        /// <summary>Returns the style left editable once the sheet is protected.</summary>
        public CellStyle Unlocked(bool value = true) =>
            style with { Protection = style.Protection with { Locked = !value } };

        /// <summary>Returns the style whose formula the formula bar hides on a protected sheet.</summary>
        public CellStyle HideFormula(bool value = true) =>
            style with { Protection = style.Protection with { Hidden = value } };
    }
}
