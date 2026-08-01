namespace Cellwright.Formulas.Functions;

/// <summary>
/// Finds the rate at which a stream of cash flows balances. <c>RATE</c> and <c>IRR</c> differ
/// only in the function they hand over; neither can be rearranged into a formula, so both are
/// searched for.
/// </summary>
internal static class RateSolver
{
    /// <summary>How close to zero the balance has to come to count as solved.</summary>
    private const double Tolerance = 1e-10;

    /// <summary>
    /// Searches from a guess. Newton reaches the root in a handful of steps whenever the guess
    /// is anywhere near it, but it wanders off on flat or badly scaled problems, so a failed
    /// run falls back to bracketing the root and halving the interval — slower, and it cannot
    /// miss a root it has surrounded.
    /// </summary>
    /// <param name="balance">What the flows come to at a given rate.</param>
    /// <param name="guess">Where to start; Excel's own default is a tenth.</param>
    public static FormulaValue Solve(Func<double, double> balance, double guess)
    {
        double rate = guess;
        for (int step = 0; step < 64; step++)
        {
            double value = balance(rate);
            if (double.IsNaN(value) || double.IsInfinity(value))
                break;
            if (Math.Abs(value) < Tolerance)
                return FormulaValue.Number(rate);

            // A numeric slope keeps the caller free to hand over any function it likes, and the
            // derivatives of these two are not worth spelling out for the few steps this takes.
            double delta = Math.Max(Math.Abs(rate) * 1e-6, 1e-8);
            double slope = (balance(rate + delta) - balance(rate - delta)) / (2 * delta);
            if (slope == 0 || double.IsNaN(slope) || double.IsInfinity(slope))
                break;

            double next = rate - (value / slope);
            if (double.IsNaN(next) || double.IsInfinity(next) || next <= -1)
                break;
            if (Math.Abs(next - rate) < Tolerance)
                return FormulaValue.Number(next);

            rate = next;
        }

        return Bracket(balance);
    }

    /// <summary>
    /// Walks outward looking for two rates the balance has opposite signs at. Without such a
    /// pair there is nothing to close in on, and the answer is that no rate exists.
    /// </summary>
    private static FormulaValue Bracket(Func<double, double> balance)
    {
        double previous = Probe(0);
        double previousValue = balance(previous);
        for (int step = 1; step <= Probes; step++)
        {
            double current = Probe(step);
            double value = balance(current);
            if (Usable(value) && Usable(previousValue) && Math.Sign(value) != Math.Sign(previousValue))
                return Halve(balance, previous, current);

            previous = current;
            previousValue = value;
        }

        return FormulaValue.Error(ExcelError.Number);
    }

    private const int Probes = 300;

    /// <summary>
    /// Rates the scan visits: every percent across the band a rate realistically falls in, then
    /// ever wider steps for the runaway returns a short venture can show.
    /// </summary>
    private static double Probe(int step) =>
        step <= 200 ? -0.999999 + (step * 0.01) : Math.Pow(1.2, step - 200);

    private static bool Usable(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    /// <summary>Halves an interval the root is known to lie in until the two ends meet.</summary>
    private static FormulaValue Halve(Func<double, double> balance, double low, double high)
    {
        double lowValue = balance(low);
        for (int step = 0; step < 200; step++)
        {
            double middle = (low + high) / 2;
            double value = balance(middle);
            if (Math.Abs(value) < Tolerance || Math.Abs(high - low) < 1e-14)
                return FormulaValue.Number(middle);

            if (Math.Sign(value) == Math.Sign(lowValue))
            {
                low = middle;
                lowValue = value;
            }
            else
            {
                high = middle;
            }
        }

        return FormulaValue.Number((low + high) / 2);
    }
}
