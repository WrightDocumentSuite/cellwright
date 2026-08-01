namespace Cellwright.Formulas.Functions;

/// <summary>
/// The four Bessel families. Small arguments come from the defining series — computed, not
/// tabulated — and large ones from Hankel's asymptotic expansion, whose terms are likewise
/// generated on the fly; the crossover sits where both give twelve good digits.
/// </summary>
internal static class BesselFunctions
{
    private const double EulerGamma = 0.57721566490153286;

    /// <summary>Where the series hands over to the asymptotic expansion.</summary>
    private const double Crossover = 12;

    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("BESSELJ", 2, 2, (c, a) => Bessel(c, a, Kind.J));
        builder.Add("BESSELY", 2, 2, (c, a) => Bessel(c, a, Kind.Y));
        builder.Add("BESSELI", 2, 2, (c, a) => Bessel(c, a, Kind.I));
        builder.Add("BESSELK", 2, 2, (c, a) => Bessel(c, a, Kind.K));
    }

    private enum Kind
    {
        J,
        Y,
        I,
        K,
    }

    private static FormulaValue Bessel(EvalContext context, ReadOnlySpan<FormulaValue> args, Kind kind)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double orderRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int order = (int)Math.Truncate(orderRaw);
        if (order < 0)
            return FormulaValue.Error(ExcelError.Number);

        // J and I take a negative argument by symmetry; the second-kind pair does not exist there.
        double sign = 1;
        if (x < 0)
        {
            if (kind is Kind.Y or Kind.K)
                return FormulaValue.Error(ExcelError.Number);
            x = -x;
            if (kind == Kind.J && order % 2 == 1)
                sign = -1;
        }

        double result = kind switch
        {
            Kind.J => J(order, x),
            Kind.Y => Y(order, x),
            Kind.I => I(order, x),
            _ => K(order, x),
        };

        result *= sign;
        return double.IsNaN(result) || double.IsInfinity(result)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(result);
    }

    private static double J(int n, double x)
    {
        if (x == 0)
            return n == 0 ? 1 : 0;
        if (x > Crossover)
            return OscillatoryAsymptotic(n, x, second: false);

        // Σ (-1)ᵏ (x/2)^(n+2k) / (k! (n+k)!)
        double half = x / 2;
        double term = Math.Exp((n * Math.Log(half)) - SpecialFunctions.LogGamma(n + 1));
        double sum = term;
        for (int k = 1; k <= 200; k++)
        {
            term *= -(half * half) / (k * (double)(n + k));
            sum += term;
            if (Math.Abs(term) < Math.Abs(sum) * 1e-17 + 1e-300)
                break;
        }

        return sum;
    }

    private static double I(int n, double x)
    {
        if (x == 0)
            return n == 0 ? 1 : 0;

        // The same series as J with every sign positive; it converges for any argument and
        // never cancels, so no asymptotic branch is needed until the double overflows.
        double half = x / 2;
        double term = Math.Exp((n * Math.Log(half)) - SpecialFunctions.LogGamma(n + 1));
        double sum = term;
        for (int k = 1; k <= 400; k++)
        {
            term *= half * half / (k * (double)(n + k));
            sum += term;
            if (term < sum * 1e-17)
                break;
        }

        return sum;
    }

    private static double Y(int n, double x)
    {
        if (x == 0)
            return double.NegativeInfinity;
        if (x > Crossover)
            return OscillatoryAsymptotic(n, x, second: true);

        // Y_n = (2/π)(ln(x/2)+γ)·J_n − (1/π)·Σ_{k<n}((n−k−1)!/k!)(x/2)^{2k−n}
        //       − (1/π)·Σ_k (−1)ᵏ(H_k+H_{n+k})(x/2)^{n+2k}/(k!(n+k)!)
        double half = x / 2;
        double result = 2 / Math.PI * (Math.Log(half) + EulerGamma) * J(n, x);
        result -= FiniteSum(n, half, alternating: false) / Math.PI;

        double harmonicK = 0;
        double harmonicNk = Harmonic(n);
        double term = Math.Exp((n * Math.Log(half)) - SpecialFunctions.LogGamma(n + 1));
        double sum = term * (harmonicK + harmonicNk);
        for (int k = 1; k <= 200; k++)
        {
            term *= -(half * half) / (k * (double)(n + k));
            harmonicK += 1.0 / k;
            harmonicNk += 1.0 / (n + k);
            double piece = term * (harmonicK + harmonicNk);
            sum += piece;
            if (Math.Abs(piece) < (Math.Abs(sum) + Math.Abs(result)) * 1e-17 + 1e-300)
                break;
        }

        return result - (sum / Math.PI);
    }

    private static double K(int n, double x)
    {
        if (x == 0)
            return double.PositiveInfinity;
        if (x > 9)
            return DecayingAsymptotic(n, x);

        // K_n = ½ Σ_{k<n}((n−k−1)!/k!)(−1)ᵏ(x/2)^{2k−n} + (−1)^{n+1} ln(x/2)·I_n
        //       + (−1)ⁿ ½ Σ_k (H_k+H_{n+k})(x/2)^{n+2k}/(k!(n+k)!)
        double half = x / 2;
        double quarter = half * half;
        double sign = n % 2 == 0 ? 1 : -1;

        // Euler's constant folds into the logarithm the same way it does for Y: the digamma
        // terms of the infinite series each carry −γ, and their sum against 1/(k!(n+k)!) is
        // exactly the I series.
        double result = FiniteSum(n, half, alternating: true) / 2;
        result += sign * -1 * (Math.Log(half) + EulerGamma) * I(n, x);

        double harmonicK = 0;
        double harmonicNk = Harmonic(n);
        double term = Math.Exp((n * Math.Log(half)) - SpecialFunctions.LogGamma(n + 1));
        double sum = term * (harmonicK + harmonicNk);
        for (int k = 1; k <= 400; k++)
        {
            term *= quarter / (k * (double)(n + k));
            harmonicK += 1.0 / k;
            harmonicNk += 1.0 / (n + k);
            double piece = term * (harmonicK + harmonicNk);
            sum += piece;
            if (piece < (Math.Abs(sum) + 1e-300) * 1e-17)
                break;
        }

        return result + (sign * sum / 2);
    }

    private static double Harmonic(int n)
    {
        double sum = 0;
        for (int i = 1; i <= n; i++)
            sum += 1.0 / i;
        return sum;
    }

    /// <summary>
    /// The finite head both second-kind series share: Σ_{k&lt;n} ((n−k−1)!/k!)·(±1)ᵏ·h^{2k−n}.
    /// Orders are small, so each term is taken straight from the log-gamma rather than by ratio.
    /// </summary>
    private static double FiniteSum(int n, double half, bool alternating)
    {
        double sum = 0;
        for (int k = 0; k < n; k++)
        {
            double term = Math.Exp(
                SpecialFunctions.LogGamma(n - k) - SpecialFunctions.LogGamma(k + 1) +
                (((2.0 * k) - n) * Math.Log(half)));
            sum += alternating && k % 2 == 1 ? -term : term;
        }

        return sum;
    }

    /// <summary>
    /// Hankel's expansion for the oscillatory pair: J uses cos·P − sin·Q, Y uses sin·P + cos·Q,
    /// around the phase ω = x − nπ/2 − π/4. The P and Q sums stop at their smallest term, which
    /// is how an asymptotic series is meant to be read.
    /// </summary>
    private static double OscillatoryAsymptotic(int n, double x, bool second)
    {
        (double p, double q) = HankelSums(n, x);
        double omega = x - (n * Math.PI / 2) - (Math.PI / 4);
        double amplitude = Math.Sqrt(2 / (Math.PI * x));
        return second
            ? amplitude * ((Math.Sin(omega) * p) + (Math.Cos(omega) * q))
            : amplitude * ((Math.Cos(omega) * p) - (Math.Sin(omega) * q));
    }

    private static (double P, double Q) HankelSums(int n, double x)
    {
        double mu = 4.0 * n * n;
        double p = 1;
        double q = 0;
        double term = 1;
        double eight = 8 * x;
        double previous = double.MaxValue;
        for (int k = 1; k <= 40; k++)
        {
            double odd = (2.0 * k) - 1;
            term *= (mu - (odd * odd)) / (k * eight);
            if (Math.Abs(term) >= previous)
                break;

            previous = Math.Abs(term);
            if (k % 4 == 1)
                q += term;
            else if (k % 4 == 2)
                p -= term;
            else if (k % 4 == 3)
                q -= term;
            else
                p += term;

            if (Math.Abs(term) < 1e-17)
                break;
        }

        return (p, q);
    }

    /// <summary>The decaying expansion for K: √(π/2x)·e⁻ˣ·Σ (4n²−1²)(4n²−3²)…/(k!(8x)ᵏ).</summary>
    private static double DecayingAsymptotic(int n, double x)
    {
        double mu = 4.0 * n * n;
        double sum = 1;
        double term = 1;
        double previous = double.MaxValue;
        for (int k = 1; k <= 40; k++)
        {
            double odd = (2.0 * k) - 1;
            term *= (mu - (odd * odd)) / (k * 8 * x);
            if (Math.Abs(term) >= previous || Math.Abs(term) < 1e-18)
            {
                sum += term;
                break;
            }

            previous = Math.Abs(term);
            sum += term;
        }

        return Math.Sqrt(Math.PI / (2 * x)) * Math.Exp(-x) * sum;
    }
}
