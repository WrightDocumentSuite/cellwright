using System.Globalization;
using System.Text;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// String functions. Positions are 1-based as in Excel, and results longer than a cell can
/// hold become <c>#VALUE!</c>.
/// </summary>
internal static class TextFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("CONCAT", 1, FunctionDefinition.Unlimited, Concat, wantsRawReferences: true);
        builder.Add("CONCATENATE", 1, FunctionDefinition.Unlimited, Concatenate);
        builder.Add("TEXTJOIN", 3, FunctionDefinition.Unlimited, TextJoin, wantsRawReferences: true);
        builder.Add("LEN", 1, 1, (c, a) => FormulaValue.Number(Args.Text(c, a[0]).Length));
        builder.Add("LEFT", 1, 2, (c, a) => Slice(c, a, fromLeft: true));
        builder.Add("RIGHT", 1, 2, (c, a) => Slice(c, a, fromLeft: false));
        builder.Add("MID", 3, 3, Mid);
        builder.Add("TRIM", 1, 1, (c, a) => Wrap(Trim(Args.Text(c, a[0]))));
        builder.Add("CLEAN", 1, 1, (c, a) => Wrap(Clean(Args.Text(c, a[0]))));
        builder.Add("UPPER", 1, 1, (c, a) => Wrap(Args.Text(c, a[0]).ToUpperInvariant()));
        builder.Add("LOWER", 1, 1, (c, a) => Wrap(Args.Text(c, a[0]).ToLowerInvariant()));
        builder.Add("PROPER", 1, 1, (c, a) => Wrap(Proper(Args.Text(c, a[0]))));
        builder.Add("SUBSTITUTE", 3, 4, Substitute);
        builder.Add("REPLACE", 4, 4, Replace);
        builder.Add("FIND", 2, 3, (c, a) => Search(c, a, StringComparison.Ordinal, wildcards: false));
        builder.Add("SEARCH", 2, 3, (c, a) => Search(c, a, StringComparison.OrdinalIgnoreCase, wildcards: true));
        builder.Add("REPT", 2, 2, Repeat);
        builder.Add("EXACT", 2, 2, (c, a) => FormulaValue.Boolean(string.Equals(Args.Text(c, a[0]), Args.Text(c, a[1]), StringComparison.Ordinal)));
        builder.Add("CHAR", 1, 1, Char);
        builder.Add("CODE", 1, 1, Code);
        builder.Add("VALUE", 1, 1, Value);
        builder.Add("T", 1, 1, (c, a) => Operators.Dereference(a[0], c) is { IsText: true } text ? FormulaValue.Scalar(text) : FormulaValue.Text(""));
        builder.Add("TEXT", 2, 2, Text);
        builder.Add("FIXED", 1, 3, Fixed);
    }

    /// <summary>
    /// <c>FIXED(number, [decimals], [no_commas])</c> rounds a number and returns it as text.
    /// Negative decimals round to the left of the point, so <c>FIXED(12345.6,-2)</c> is 12,300.
    /// </summary>
    private static FormulaValue Fixed(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        int decimals = 2;
        if (args.Length > 1)
        {
            if (!Args.TryNumber(context, args[1], out double requested, out error))
                return FormulaValue.Scalar(error);
            decimals = (int)Math.Truncate(requested);
        }

        bool grouped = true;
        if (args.Length > 2)
        {
            if (!Args.TryBoolean(context, args[2], out bool suppress, out error))
                return FormulaValue.Scalar(error);
            grouped = !suppress;
        }

        if (decimals < 0)
        {
            double scale = Math.Pow(10, -decimals);
            number = Math.Round(number / scale, MidpointRounding.AwayFromZero) * scale;
            decimals = 0;
        }

        decimals = Math.Min(decimals, 127);
        string format = (grouped ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : "");
        return Wrap(number.ToString(format, CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// <c>TEXT(value, format)</c> renders a value through a number format code. Formatting is
    /// culture-invariant, like the rest of the library: the separators and month names do not
    /// change with the machine the code runs on.
    /// </summary>
    private static FormulaValue Text(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue value = Operators.Dereference(args[0], context);
        if (value.IsError)
            return FormulaValue.Scalar(value);

        string code = Args.Text(context, args[1]);
        CultureInfo culture = context.FormatCulture;
        if (!ReferenceEquals(culture, CultureInfo.InvariantCulture))
            code = LocalizedFormatCode.ToInvariant(code, culture);
        return Wrap(Cellwright.Styles.NumberFormatRenderer.Format(value, code, culture, context.DateSystem));
    }

    private static FormulaValue Concat(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        var builder = new StringBuilder();
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsError)
                    return FormulaValue.Scalar(value);
                builder.Append(Coercion.ToText(value, context.DateSystem));
            }
        }

        return Wrap(builder.ToString());
    }

    private static FormulaValue Concatenate(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        var builder = new StringBuilder();
        foreach (ref readonly FormulaValue arg in args)
            builder.Append(Args.Text(context, arg));
        return Wrap(builder.ToString());
    }

    private static FormulaValue TextJoin(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string separator = Args.Text(context, args[0]);
        if (!Args.TryBoolean(context, args[1], out bool ignoreEmpty, out CellValue error))
            return FormulaValue.Scalar(error);

        var builder = new StringBuilder();
        bool first = true;
        foreach (ref readonly FormulaValue arg in args[2..])
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsError)
                    return FormulaValue.Scalar(value);
                if (ignoreEmpty && (value.IsBlank || (value.IsText && value.GetText().Length == 0)))
                    continue;

                if (!first)
                    builder.Append(separator);
                builder.Append(Coercion.ToText(value, context.DateSystem));
                first = false;
            }
        }

        return Wrap(builder.ToString());
    }

    private static FormulaValue Slice(EvalContext context, ReadOnlySpan<FormulaValue> args, bool fromLeft)
    {
        string text = Args.Text(context, args[0]);
        int count = 1;
        if (args.Length > 1 && !Args.TryInteger(context, args[1], out count, out CellValue error))
            return FormulaValue.Scalar(error);
        if (count < 0)
            return FormulaValue.Error(ExcelError.Value);

        count = Math.Min(count, text.Length);
        return Wrap(fromLeft ? text[..count] : text[^count..]);
    }

    private static FormulaValue Mid(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        if (!Args.TryInteger(context, args[1], out int start, out CellValue error) ||
            !Args.TryInteger(context, args[2], out int count, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (start < 1 || count < 0)
            return FormulaValue.Error(ExcelError.Value);
        if (start > text.Length)
            return FormulaValue.Text("");

        int available = Math.Min(count, text.Length - start + 1);
        return Wrap(text.Substring(start - 1, available));
    }

    private static FormulaValue Substitute(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        string search = Args.Text(context, args[1]);
        string replacement = Args.Text(context, args[2]);
        if (search.Length == 0)
            return Wrap(text);

        if (args.Length == 3)
            return Wrap(text.Replace(search, replacement, StringComparison.Ordinal));

        if (!Args.TryInteger(context, args[3], out int occurrence, out CellValue error))
            return FormulaValue.Scalar(error);
        if (occurrence < 1)
            return FormulaValue.Error(ExcelError.Value);

        int position = -1;
        for (int i = 0; i < occurrence; i++)
        {
            position = text.IndexOf(search, position + 1, StringComparison.Ordinal);
            if (position < 0)
                return Wrap(text);
        }

        return Wrap(string.Concat(text.AsSpan(0, position), replacement, text.AsSpan(position + search.Length)));
    }

    private static FormulaValue Replace(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        if (!Args.TryInteger(context, args[1], out int start, out CellValue error) ||
            !Args.TryInteger(context, args[2], out int count, out error))
        {
            return FormulaValue.Scalar(error);
        }

        string replacement = Args.Text(context, args[3]);
        if (start < 1 || count < 0)
            return FormulaValue.Error(ExcelError.Value);

        int prefix = Math.Min(start - 1, text.Length);
        int suffixStart = Math.Min(prefix + count, text.Length);
        return Wrap(string.Concat(text.AsSpan(0, prefix), replacement, text.AsSpan(suffixStart)));
    }

    private static FormulaValue Search(
        EvalContext context, ReadOnlySpan<FormulaValue> args, StringComparison comparison, bool wildcards)
    {
        string needle = Args.Text(context, args[0]);
        string haystack = Args.Text(context, args[1]);
        int start = 1;
        if (args.Length > 2 && !Args.TryInteger(context, args[2], out start, out CellValue error))
            return FormulaValue.Scalar(error);
        if (start < 1 || start > haystack.Length + 1)
            return FormulaValue.Error(ExcelError.Value);

        int index = wildcards && Wildcard.ContainsWildcards(needle)
            ? Wildcard.IndexOf(haystack, needle, start - 1)
            : haystack.IndexOf(needle, start - 1, comparison);

        return index < 0 ? FormulaValue.Error(ExcelError.Value) : FormulaValue.Number(index + 1);
    }

    private static FormulaValue Repeat(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        if (!Args.TryInteger(context, args[1], out int count, out CellValue error))
            return FormulaValue.Scalar(error);
        if (count < 0)
            return FormulaValue.Error(ExcelError.Value);
        if ((long)text.Length * count > SheetLimits.MaxCellTextLength)
            return FormulaValue.Error(ExcelError.Value);

        return Wrap(string.Concat(Enumerable.Repeat(text, count)));
    }

    private static FormulaValue Char(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int code, out CellValue error))
            return FormulaValue.Scalar(error);
        return code is < 1 or > 255
            ? FormulaValue.Error(ExcelError.Value)
            : FormulaValue.Text(((char)code).ToString());
    }

    private static FormulaValue Code(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        return text.Length == 0 ? FormulaValue.Error(ExcelError.Value) : FormulaValue.Number(text[0]);
    }

    private static FormulaValue Value(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue value = Operators.Dereference(args[0], context);
        if (value.IsNumber)
            return FormulaValue.Scalar(value);

        string text = Args.Text(context, args[0]).Trim();
        CellValue number = Coercion.ParseNumber(text);
        if (!number.IsError)
            return FormulaValue.Scalar(number);

        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)
            ? FormulaValue.Number(ExcelSerialDate.FromDateTime(date, context.DateSystem))
            : FormulaValue.Error(ExcelError.Value);
    }

    private static string Trim(string text)
    {
        // Excel's TRIM also collapses runs of interior spaces.
        var builder = new StringBuilder(text.Length);
        bool pendingSpace = false;
        foreach (char c in text.Trim(' '))
        {
            if (c == ' ')
            {
                pendingSpace = true;
                continue;
            }

            if (pendingSpace && builder.Length > 0)
                builder.Append(' ');
            pendingSpace = false;
            builder.Append(c);
        }

        return builder.ToString();
    }

    private static string Clean(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (!char.IsControl(c))
                builder.Append(c);
        }

        return builder.ToString();
    }

    private static string Proper(string text)
    {
        Span<char> buffer = text.Length <= 256 ? stackalloc char[text.Length] : new char[text.Length];
        bool startOfWord = true;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            buffer[i] = startOfWord ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c);
            startOfWord = !char.IsLetter(c);
        }

        return new string(buffer);
    }

    private static FormulaValue Wrap(string text) =>
        text.Length > SheetLimits.MaxCellTextLength ? FormulaValue.Error(ExcelError.Value) : FormulaValue.Text(text);
}
