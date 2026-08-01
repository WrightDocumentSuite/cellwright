namespace Cellwright;

/// <summary>How a gradient sweeps the cell, matching <c>ST_GradientType</c> (§18.18.33).</summary>
public enum GradientFillType : byte
{
    /// <summary>Color bands run perpendicular to <see cref="GradientFill.Degree"/>.</summary>
    Linear = 0,

    /// <summary>Colors radiate from the rectangle described by the four convergence insets.</summary>
    Path,
}

/// <summary>One color stop of a gradient: where it sits (0..1) and what it paints.</summary>
public readonly record struct GradientStop(double Position, ExcelColor Color);

/// <summary>
/// A gradient cell fill (§18.8.24). Two or more <see cref="Stops"/> blend into each other,
/// either along a line (<see cref="GradientFillType.Linear"/>) or outward from an inner
/// rectangle (<see cref="GradientFillType.Path"/>).
/// </summary>
public sealed record GradientFill
{
    /// <summary>The shape of the sweep.</summary>
    public GradientFillType Type { get; init; }

    /// <summary>Angle of a linear gradient in degrees, clockwise from left-to-right.</summary>
    public double Degree { get; init; }

    /// <summary>Left edge of the inner rectangle of a path gradient, as a fraction of the cell.</summary>
    public double Left { get; init; }

    /// <summary>Right edge of the inner rectangle, as a fraction of the cell.</summary>
    public double Right { get; init; }

    /// <summary>Top edge of the inner rectangle, as a fraction of the cell.</summary>
    public double Top { get; init; }

    /// <summary>Bottom edge of the inner rectangle, as a fraction of the cell.</summary>
    public double Bottom { get; init; }

    /// <summary>The color stops in file order; positions run from 0 to 1.</summary>
    public IReadOnlyList<GradientStop> Stops { get; init; } = [];

    /// <summary>A left-to-right linear gradient between two colors.</summary>
    public static GradientFill Linear(ExcelColor from, ExcelColor to, double degree = 0) => new()
    {
        Degree = degree,
        Stops = [new GradientStop(0, from), new GradientStop(1, to)],
    };

    /// <summary>
    /// Value equality including the stops, element by element. Records compare list properties
    /// by reference, and the style pool interns on equality — without this, every save of the
    /// same gradient would mint a new fill entry.
    /// </summary>
    public bool Equals(GradientFill? other)
    {
        if (other is null)
            return false;
        if (ReferenceEquals(this, other))
            return true;
        if (Type != other.Type || Degree != other.Degree ||
            Left != other.Left || Right != other.Right || Top != other.Top || Bottom != other.Bottom ||
            Stops.Count != other.Stops.Count)
            return false;

        for (int i = 0; i < Stops.Count; i++)
        {
            if (Stops[i] != other.Stops[i])
                return false;
        }

        return true;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Type);
        hash.Add(Degree);
        hash.Add(Left);
        hash.Add(Right);
        hash.Add(Top);
        hash.Add(Bottom);
        foreach (GradientStop stop in Stops)
            hash.Add(stop);
        return hash.ToHashCode();
    }
}
