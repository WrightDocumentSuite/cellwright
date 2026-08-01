using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing sheet and workbook protection. Both sides of the file format use the
/// same pair of password forms, so the hash attributes are handled once here.
/// </summary>
internal static class ProtectionPart
{
    /// <summary>Reads a <c>sheetProtection</c> element into the model.</summary>
    public static void ReadSheet(XmlReader xml, SheetProtection protection)
    {
        protection.IsEnabled = IsTrue(xml.GetAttribute("sheet"));
        protection.AllowSelectLockedCells = !IsTrue(xml.GetAttribute("selectLockedCells"));
        protection.AllowSelectUnlockedCells = !IsTrue(xml.GetAttribute("selectUnlockedCells"));
        protection.AllowFormatCells = IsTrue(xml.GetAttribute("formatCells"));
        protection.AllowFormatColumns = IsTrue(xml.GetAttribute("formatColumns"));
        protection.AllowFormatRows = IsTrue(xml.GetAttribute("formatRows"));
        protection.AllowInsertColumns = IsTrue(xml.GetAttribute("insertColumns"));
        protection.AllowInsertRows = IsTrue(xml.GetAttribute("insertRows"));
        protection.AllowInsertHyperlinks = IsTrue(xml.GetAttribute("insertHyperlinks"));
        protection.AllowDeleteColumns = IsTrue(xml.GetAttribute("deleteColumns"));
        protection.AllowDeleteRows = IsTrue(xml.GetAttribute("deleteRows"));
        protection.AllowSort = IsTrue(xml.GetAttribute("sort"));
        protection.AllowAutoFilter = IsTrue(xml.GetAttribute("autoFilter"));
        protection.AllowPivotTables = IsTrue(xml.GetAttribute("pivotTables"));
        protection.AllowEditObjects = !IsTrue(xml.GetAttribute("objects"));
        protection.AllowEditScenarios = !IsTrue(xml.GetAttribute("scenarios"));
        protection.Password = ReadPassword(xml);
    }

    /// <summary>Reads a <c>protectedRanges</c> element into the model.</summary>
    public static void ReadProtectedRanges(string element, SheetProtection protection)
    {
        using var xml = XmlReader.Create(new StringReader(element), XmlDefaults.ReaderSettings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "protectedRange")
                continue;

            string? sqref = xml.GetAttribute("sqref");
            if (string.IsNullOrWhiteSpace(sqref))
                continue;

            try
            {
                protection.EditableRanges.Add(new ProtectedRange
                {
                    Name = xml.GetAttribute("name") ?? $"Range{protection.EditableRanges.Count + 1}",
                    Ranges = DataValidation.ParseRanges(sqref),
                    Password = ReadPassword(xml),
                });
            }
            catch (FormatException)
            {
                // A range Excel cannot express either; the element itself still travels as markup.
            }
        }
    }

    /// <summary>Reads a <c>workbookProtection</c> element into the model.</summary>
    public static void ReadWorkbook(XmlReader xml, WorkbookProtection protection)
    {
        protection.LockStructure = IsTrue(xml.GetAttribute("lockStructure"));
        protection.LockWindows = IsTrue(xml.GetAttribute("lockWindows"));
        protection.Password = PasswordHash.FromFile(
            xml.GetAttribute("workbookHashValue"),
            xml.GetAttribute("workbookSaltValue"),
            xml.GetAttribute("workbookAlgorithmName"),
            SpinCount(xml.GetAttribute("workbookSpinCount")),
            xml.GetAttribute("workbookPassword"));
    }

    /// <summary>Writes <c>sheetProtection</c> and, when there are exceptions, <c>protectedRanges</c>.</summary>
    public static void WriteSheet(Utf8XmlWriter xml, SheetProtection protection)
    {
        if (!protection.IsEnabled)
            return;

        xml.WriteRaw("<sheetProtection"u8);
        WritePassword(xml, protection.Password, "");
        xml.WriteRaw(" sheet=\"1\""u8);
        WriteFlag(xml, " objects=\"1\""u8, !protection.AllowEditObjects);
        WriteFlag(xml, " scenarios=\"1\""u8, !protection.AllowEditScenarios);
        WriteFlag(xml, " formatCells=\"1\""u8, protection.AllowFormatCells);
        WriteFlag(xml, " formatColumns=\"1\""u8, protection.AllowFormatColumns);
        WriteFlag(xml, " formatRows=\"1\""u8, protection.AllowFormatRows);
        WriteFlag(xml, " insertColumns=\"1\""u8, protection.AllowInsertColumns);
        WriteFlag(xml, " insertRows=\"1\""u8, protection.AllowInsertRows);
        WriteFlag(xml, " insertHyperlinks=\"1\""u8, protection.AllowInsertHyperlinks);
        WriteFlag(xml, " deleteColumns=\"1\""u8, protection.AllowDeleteColumns);
        WriteFlag(xml, " deleteRows=\"1\""u8, protection.AllowDeleteRows);
        WriteFlag(xml, " selectLockedCells=\"1\""u8, !protection.AllowSelectLockedCells);
        WriteFlag(xml, " sort=\"1\""u8, protection.AllowSort);
        WriteFlag(xml, " autoFilter=\"1\""u8, protection.AllowAutoFilter);
        WriteFlag(xml, " pivotTables=\"1\""u8, protection.AllowPivotTables);
        WriteFlag(xml, " selectUnlockedCells=\"1\""u8, !protection.AllowSelectUnlockedCells);
        xml.WriteRaw("/>"u8);

        if (protection.EditableRanges.Count == 0)
            return;

        xml.WriteRaw("<protectedRanges>"u8);
        foreach (ProtectedRange range in protection.EditableRanges)
        {
            xml.WriteRaw("<protectedRange"u8);
            WritePassword(xml, range.Password, "");
            xml.WriteRaw(" sqref=\""u8);
            for (int i = 0; i < range.Ranges.Count; i++)
            {
                if (i > 0)
                    xml.WriteRaw(" "u8);
                xml.WriteFormattable(range.Ranges[i]);
            }

            xml.WriteRaw("\" name=\""u8);
            xml.WriteAttributeText(range.Name);
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</protectedRanges>"u8);
    }

    /// <summary>Writes <c>workbookProtection</c>, or nothing when the workbook is open.</summary>
    public static void WriteWorkbook(Utf8XmlWriter xml, WorkbookProtection protection)
    {
        if (!protection.IsEnabled)
            return;

        xml.WriteRaw("<workbookProtection"u8);
        WritePassword(xml, protection.Password, "workbook");
        WriteFlag(xml, " lockStructure=\"1\""u8, protection.LockStructure);
        WriteFlag(xml, " lockWindows=\"1\""u8, protection.LockWindows);
        xml.WriteRaw("/>"u8);
    }

    private static PasswordHash? ReadPassword(XmlReader xml) => PasswordHash.FromFile(
        xml.GetAttribute("hashValue"),
        xml.GetAttribute("saltValue"),
        xml.GetAttribute("algorithmName"),
        SpinCount(xml.GetAttribute("spinCount")),
        xml.GetAttribute("password"));

    /// <summary>
    /// Writes the password attributes. The prefix distinguishes the workbook element, whose
    /// attributes carry the same four names behind a <c>workbook</c> prefix.
    /// </summary>
    private static void WritePassword(Utf8XmlWriter xml, PasswordHash? password, string prefix)
    {
        if (password is null)
            return;

        if (password.IsLegacy)
        {
            WriteAttribute(xml, prefix, "password", password.Value);
            return;
        }

        WriteAttribute(xml, prefix, "algorithmName", password.Algorithm ?? "SHA-512");
        WriteAttribute(xml, prefix, "hashValue", password.Value);
        WriteAttribute(xml, prefix, "saltValue", password.Salt ?? "");
        WriteAttribute(xml, prefix, "spinCount", password.SpinCount.ToString(CultureInfo.InvariantCulture));
    }

    private static void WriteAttribute(Utf8XmlWriter xml, string prefix, string name, string value)
    {
        xml.WriteRaw(" "u8);
        xml.WriteRawXml(prefix.Length == 0 ? name : prefix + char.ToUpperInvariant(name[0]) + name[1..]);
        xml.WriteRaw("=\""u8);
        xml.WriteAttributeText(value);
        xml.WriteRaw("\""u8);
    }

    private static void WriteFlag(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> attribute, bool value)
    {
        if (value)
            xml.WriteRaw(attribute);
    }

    private static int SpinCount(string? text) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) ? value : 0;

    private static bool IsTrue(string? value) => value is "1" or "true" or "TRUE";
}
