namespace Cellwright.Formulas.Functions;

/// <summary>
/// The handful of mathematical functions every statistical distribution is built out of: the
/// gamma function and its two incomplete forms, the incomplete beta, and the error function.
/// </summary>
/// <remarks>
/// Nothing here is a worksheet function. Between them these four cover every continuous
/// distribution Excel offers — the normal is the error function, chi-squared and Poisson are
/// the incomplete gamma, and the t, F and binomial families are the incomplete beta — so the
/// accuracy of the whole catalog rests on these.
/// </remarks>
internal static class SpecialFunctions
{
    /// <summary>Where a series or a continued fraction is close enough to have converged.</summary>
    private const double Tolerance = 1e-15;

    private const int MaxIterations = 300;

    /// <summary>
    /// Natural logarithm of the gamma function, by the Lanczos approximation. The logarithm
    /// rather than the value itself because the value overflows a double past about 170, while
    /// every use of it here is a ratio that stays small.
    /// </summary>
    public static double LogGamma(double x)
    {
        if (x <= 0)
            return double.NaN;

        // The approximation is stated for the right half-plane; the reflection formula covers
        // the rest, though nothing in this catalog asks for it.
        double sum = Lanczos[0];
        for (int i = 1; i < Lanczos.Length; i++)
            sum += Lanczos[i] / (x + i - 1);

        double t = x + Lanczos.Length - 2.5;
        return (0.5 * Math.Log(2 * Math.PI)) + ((x - 0.5) * Math.Log(t)) - t + Math.Log(sum);
    }

    private static readonly double[] Lanczos =
    [
        0.99999999999980993, 676.5203681218851, -1259.1392167224028,
        771.32342877765313, -176.61502916214059, 12.507343278686905,
        -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
    ];

    /// <summary>Gamma itself, for the arguments small enough to have one.</summary>
    public static double Gamma(double x)
    {
        if (x <= 0 && x == Math.Floor(x))
            return double.NaN;
        if (x < 0)
            return Math.PI / (Math.Sin(Math.PI * x) * Math.Exp(LogGamma(1 - x)));

        return Math.Exp(LogGamma(x));
    }

    /// <summary>
    /// The regularized lower incomplete gamma, <c>P(shape, x)</c>: the share of a gamma
    /// distribution lying below <paramref name="x"/>.
    /// </summary>
    /// <remarks>
    /// Two expansions cover it between them. The series converges quickly while x sits below
    /// the shape and slowly beyond it; the continued fraction for the upper tail behaves the
    /// other way round, so each is used where the other would crawl.
    /// </remarks>
    public static double LowerGamma(double shape, double x)
    {
        if (shape <= 0 || x < 0)
            return double.NaN;
        if (x == 0)
            return 0;

        return x < shape + 1 ? GammaSeries(shape, x) : 1 - GammaFraction(shape, x);
    }

    /// <summary>
    /// The regularized upper incomplete gamma. Computed from its own expansion rather than by
    /// subtracting the lower one from one, which would throw away every digit of a small tail.
    /// </summary>
    public static double UpperGamma(double shape, double x)
    {
        if (shape <= 0 || x < 0)
            return double.NaN;
        if (x == 0)
            return 1;

        return x < shape + 1 ? 1 - GammaSeries(shape, x) : GammaFraction(shape, x);
    }

    private static double GammaSeries(double shape, double x)
    {
        double term = 1.0 / shape;
        double sum = term;
        for (int i = 1; i < MaxIterations; i++)
        {
            term *= x / (shape + i);
            sum += term;
            if (Math.Abs(term) < Math.Abs(sum) * Tolerance)
                break;
        }

        return sum * Math.Exp(-x + (shape * Math.Log(x)) - LogGamma(shape));
    }

    /// <summary>The upper tail by the modified Lentz method, which evaluates the fraction forwards.</summary>
    private static double GammaFraction(double shape, double x)
    {
        const double Tiny = 1e-300;
        double b = x + 1 - shape;
        double c = 1 / Tiny;
        double d = 1 / b;
        double result = d;

        for (int i = 1; i < MaxIterations; i++)
        {
            double a = -i * (i - shape);
            b += 2;
            d = (a * d) + b;
            if (Math.Abs(d) < Tiny)
                d = Tiny;
            c = b + (a / c);
            if (Math.Abs(c) < Tiny)
                c = Tiny;
            d = 1 / d;

            double step = d * c;
            result *= step;
            if (Math.Abs(step - 1) < Tolerance)
                break;
        }

        return result * Math.Exp(-x + (shape * Math.Log(x)) - LogGamma(shape));
    }

    /// <summary>
    /// The regularized incomplete beta, <c>I_x(a, b)</c>. The continued fraction converges
    /// briskly on one side of the distribution's centre, so the far side is reached by
    /// symmetry rather than by pushing the fraction where it struggles.
    /// </summary>
    public static double IncompleteBeta(double a, double b, double x)
    {
        if (a <= 0 || b <= 0 || x < 0 || x > 1)
            return double.NaN;
        if (x == 0 || x == 1)
            return x;

        double front = Math.Exp(
            LogGamma(a + b) - LogGamma(a) - LogGamma(b) + (a * Math.Log(x)) + (b * Math.Log(1 - x)));

        return x < (a + 1) / (a + b + 2)
            ? front * BetaFraction(a, b, x) / a
            : 1 - (Math.Exp(
                LogGamma(a + b) - LogGamma(a) - LogGamma(b) + (b * Math.Log(1 - x)) + (a * Math.Log(x)))
                * BetaFraction(b, a, 1 - x) / b);
    }

    private static double BetaFraction(double a, double b, double x)
    {
        const double Tiny = 1e-300;
        double c = 1;
        double d = 1 - ((a + b) * x / (a + 1));
        if (Math.Abs(d) < Tiny)
            d = Tiny;
        d = 1 / d;
        double result = d;

        for (int i = 1; i <= MaxIterations; i++)
        {
            // The fraction's terms alternate in form, one for each half of a step.
            double even = i * (b - i) * x / ((a + (2 * i) - 1) * (a + (2 * i)));
            result *= Advance(even, ref c, ref d);

            double odd = -(a + i) * (a + b + i) * x / ((a + (2 * i)) * (a + (2 * i) + 1));
            double step = Advance(odd, ref c, ref d);
            result *= step;
            if (Math.Abs(step - 1) < Tolerance)
                break;
        }

        return result;
    }

    private static double Advance(double numerator, ref double c, ref double d)
    {
        const double Tiny = 1e-300;
        d = 1 + (numerator * d);
        if (Math.Abs(d) < Tiny)
            d = Tiny;
        d = 1 / d;

        c = 1 + (numerator / c);
        if (Math.Abs(c) < Tiny)
            c = Tiny;

        return c * d;
    }

    /// <summary>The standard normal distribution below <paramref name="z"/>.</summary>
    public static double NormalBelow(double z) => 0.5 * Erfc(-z / Math.Sqrt(2));

    /// <summary>The height of the standard normal curve at <paramref name="z"/>.</summary>
    public static double NormalDensity(double z) => Math.Exp(-0.5 * z * z) / Math.Sqrt(2 * Math.PI);

    /// <summary>
    /// The error function. It is the incomplete gamma at a half in disguise, and taking it from
    /// there rather than from a fitted polynomial keeps the whole normal family accurate to the
    /// last digit a double holds.
    /// </summary>
    public static double Erf(double x)
    {
        double magnitude = LowerGamma(0.5, x * x);
        return x >= 0 ? magnitude : -magnitude;
    }

    /// <summary>
    /// The complementary error function. The far tail is taken from the upper expansion, where
    /// subtracting the lower one from a whole would have left nothing but rounding.
    /// </summary>
    public static double Erfc(double x) =>
        x >= 0 ? UpperGamma(0.5, x * x) : 1 + LowerGamma(0.5, x * x);

    /// <summary>
    /// The point below which a given share of a standard normal distribution lies, by the
    /// rational approximation of Wichura's AS 241 — accurate to about sixteen digits, which is
    /// as much as a double holds.
    /// </summary>
    public static double NormalQuantile(double probability)
    {
        if (probability is <= 0 or >= 1)
            return double.NaN;

        double q = probability - 0.5;
        if (Math.Abs(q) <= 0.425)
        {
            double r = 0.180625 - (q * q);
            return q * Polynomial(r, CentreNumerator) / Polynomial(r, CentreDenominator);
        }

        double tail = q < 0 ? probability : 1 - probability;
        double s = Math.Sqrt(-Math.Log(tail));
        double value = s <= 5
            ? Polynomial(s - 1.6, NearNumerator) / Polynomial(s - 1.6, NearDenominator)
            : Polynomial(s - 5, FarNumerator) / Polynomial(s - 5, FarDenominator);

        return q < 0 ? -value : value;
    }

    private static double Polynomial(double x, double[] coefficients)
    {
        double result = coefficients[^1];
        for (int i = coefficients.Length - 2; i >= 0; i--)
            result = (result * x) + coefficients[i];

        return result;
    }

    // Every array below runs from the constant term upwards, which is the order Polynomial
    // reads them in.
    private static readonly double[] CentreNumerator =
    [
        3.387132872796366608, 133.14166789178437745, 1971.5909503065514427, 13731.693765509461125,
        45921.953931549871457, 67265.770927008700853, 33430.575583588128105, 2509.0809287301226727,
    ];

    private static readonly double[] CentreDenominator =
    [
        1.0, 42.313330701600911252, 687.1870074920579083, 5394.1960214247511077,
        21213.794301586595867, 39307.89580009271061, 28729.085735721942674, 5226.495278852545925,
    ];

    private static readonly double[] NearNumerator =
    [
        1.42343711074968357734, 4.6303378461565452959, 5.7694972214606914055,
        3.64784832476320460504, 1.27045825245236838258, 0.24178072517745061177,
        0.0227238449892691845833, 7.7454501427834140764e-4,
    ];

    private static readonly double[] NearDenominator =
    [
        1.0, 2.05319162663775882187, 1.6763848301838038494, 0.68976733498510000455,
        0.14810397642748007459, 0.0151986665636164571966, 5.475938084995344946e-4,
        1.05075007164441684324e-9,
    ];

    private static readonly double[] FarNumerator =
    [
        6.6579046435011037772, 5.4637849111641143699, 1.7848265399172913358,
        0.29656057182850489123, 0.026532189526576123093, 0.0012426609473880784386,
        2.71155556874348757815e-5, 2.01033439929228813265e-7,
    ];

    private static readonly double[] FarDenominator =
    [
        1.0, 0.59983220655588793769, 0.13692988092273580531, 0.0148753612908506148525,
        7.868691311456132591e-4, 1.8463183175100546818e-5, 1.4215117583164458887e-7,
        2.04426310338993978564e-15,
    ];

    /// <summary>
    /// Finds the point at which a rising distribution reaches a given share, by bracketing it
    /// and halving. Every family without a formula of its own is inverted this way.
    /// </summary>
    /// <param name="below">The share of the distribution lying below a point.</param>
    /// <param name="probability">The share to find the point for.</param>
    /// <param name="low">A point known to sit at or below the answer.</param>
    /// <param name="high">A point to start expanding the search from.</param>
    public static double Invert(Func<double, double> below, double probability, double low, double high)
    {
        if (probability is <= 0 or >= 1)
            return double.NaN;

        // The upper end is a guess, so it is pushed out until it really is above the answer.
        for (int i = 0; i < 200 && below(high) < probability; i++)
            high = low + ((high - low) * 2);

        for (int i = 0; i < 200; i++)
        {
            double middle = (low + high) / 2;
            if (below(middle) < probability)
                low = middle;
            else
                high = middle;

            if (high - low < Math.Max(Math.Abs(high), 1) * 1e-14)
                break;
        }

        return (low + high) / 2;
    }
}
