namespace Cellwright.Formulas.Functions;

/// <summary>
/// The annuity family and the two discounted cash flow functions.
/// </summary>
/// <remarks>
/// Seven of the nine are rearrangements of a single identity — see <see cref="Balance"/> — so
/// they share its wording rather than each spelling out its own compounding. The two that
/// cannot be rearranged, <c>RATE</c> and <c>IRR</c>, hand the same identity to
/// <see cref="RateSolver"/> and let it search.
/// </remarks>
internal static class FinancialFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("PMT", 3, 5, Pmt);
        builder.Add("PV", 3, 5, Pv);
        builder.Add("FV", 3, 5, Fv);
        builder.Add("NPER", 3, 5, Nper);
        builder.Add("IPMT", 4, 6, (c, a) => Part(c, a, principal: false));
        builder.Add("PPMT", 4, 6, (c, a) => Part(c, a, principal: true));
        builder.Add("RATE", 3, 6, Rate);
        builder.Add("NPV", 2, FunctionDefinition.Unlimited, Npv, wantsRawReferences: true);
        builder.Add("IRR", 1, 2, Irr, wantsRawReferences: true);
    }

    /// <summary>
    /// What the account holds after <paramref name="nper"/> periods, before the sign is flipped
    /// to the answer's point of view: <c>pv·(1+r)^n + pmt·(1+r·type)·((1+r)^n − 1)/r</c>. At a
    /// rate of zero the compounding vanishes and the money simply piles up. A nonzero
    /// <c>type</c> says each payment falls at the start of its period rather than the end.
    /// </summary>
    private static double Balance(double rate, double nper, double pmt, double pv, double type)
    {
        if (rate == 0)
            return pv + (pmt * nper);

        double growth = Math.Pow(1 + rate, nper);
        return (pv * growth) + (pmt * Annuity(rate, growth, type));
    }

    /// <summary>Worth of a stream of one-unit payments over the same periods.</summary>
    private static double Annuity(double rate, double growth, double type) =>
        (1 + (rate * type)) * (growth - 1) / rate;

    private static FormulaValue Pmt(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> a = stackalloc double[5];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        return Payment(a[0], a[1], a[2], a[3], a[4]);
    }

    /// <summary>Solves the identity for the payment: <c>PMT(rate, nper, pv, fv, type)</c>.</summary>
    private static FormulaValue Payment(double rate, double nper, double pv, double fv, double type)
    {
        // A loan over no periods is a division by zero, and that is what Excel reports rather
        // than #NUM! (MS-OI29500 2.1.1016).
        if (nper == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);
        if (rate == 0)
            return Finite(-(pv + fv) / nper);

        double growth = Math.Pow(1 + rate, nper);
        double annuity = Annuity(rate, growth, type);
        return annuity == 0
            ? FormulaValue.Error(ExcelError.DivisionByZero)
            : Finite(-((pv * growth) + fv) / annuity);
    }

    private static FormulaValue Pv(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> a = stackalloc double[5];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        double rate = a[0], nper = a[1], pmt = a[2], fv = a[3], type = a[4];
        if (rate == 0)
            return Finite(-(fv + (pmt * nper)));

        double growth = Math.Pow(1 + rate, nper);
        return growth == 0
            ? FormulaValue.Error(ExcelError.DivisionByZero)
            : Finite(-(fv + (pmt * Annuity(rate, growth, type))) / growth);
    }

    private static FormulaValue Fv(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> a = stackalloc double[5];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        return Finite(-Balance(a[0], a[1], a[2], a[3], a[4]));
    }

    /// <summary>
    /// <c>NPER(rate, pmt, pv, fv, type)</c>. Collecting the compounding on one side turns the
    /// identity into <c>(1+r)^n = (a − fv)/(pv + a)</c>, which a logarithm undoes; a ratio at or
    /// below zero means the payments never reach the target and there is no such term.
    /// </summary>
    private static FormulaValue Nper(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> a = stackalloc double[5];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        double rate = a[0], pmt = a[1], pv = a[2], fv = a[3], type = a[4];
        if (rate == 0)
            return pmt == 0 ? FormulaValue.Error(ExcelError.Number) : Finite(-(pv + fv) / pmt);

        double annuity = pmt * (1 + (rate * type)) / rate;
        double denominator = pv + annuity;
        if (denominator == 0)
            return FormulaValue.Error(ExcelError.Number);

        double ratio = (annuity - fv) / denominator;
        return ratio <= 0 ? FormulaValue.Error(ExcelError.Number) : Finite(Math.Log(ratio) / Math.Log(1 + rate));
    }

    /// <summary>
    /// <c>IPMT</c> and <c>PPMT</c>: how one payment splits between interest and principal.
    /// </summary>
    private static FormulaValue Part(EvalContext context, ReadOnlySpan<FormulaValue> args, bool principal)
    {
        Span<double> a = stackalloc double[6];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        double rate = a[0], per = a[1], nper = a[2], pv = a[3], fv = a[4], type = a[5];
        if (per < 1 || per > nper)
            return FormulaValue.Error(ExcelError.Number);

        FormulaValue payment = Payment(rate, nper, pv, fv, type);
        if (payment.AsScalar.IsError)
            return payment;

        double pmt = payment.AsScalar.GetNumber();
        double interest = Interest(rate, per, pmt, pv, type);
        return Finite(principal ? pmt - interest : interest);
    }

    /// <summary>
    /// Interest of one period: what the balance left by the periods before it earns. A payment
    /// made at the start of its period earns nothing in the very first one, and every later one
    /// is discounted by a period against the same figure taken at the end.
    /// </summary>
    private static double Interest(double rate, double per, double pmt, double pv, double type)
    {
        if (type != 0 && per == 1)
            return 0;

        double interest = -Balance(rate, per - 1, pmt, pv, type) * rate;
        return type != 0 ? interest / (1 + rate) : interest;
    }

    private static FormulaValue Rate(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> a = stackalloc double[5];
        if (!ReadNumbers(context, args, a, out CellValue error))
            return FormulaValue.Scalar(error);

        double guess = 0.1;
        if (args.Length > 5 && !args[5].IsBlank && !Args.TryNumber(context, args[5], out guess, out error))
            return FormulaValue.Scalar(error);

        double nper = a[0], pmt = a[1], pv = a[2], fv = a[3], type = a[4];
        if (nper <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return RateSolver.Solve(rate => Balance(rate, nper, pmt, pv, type) + fv, guess);
    }

    /// <summary>
    /// <c>NPV(rate, …)</c> discounts each amount by the period it falls in, counting from one.
    /// Values a range cannot offer a number for are passed over rather than counted as zero, so
    /// a gap in the column does not silently push every later amount a period further out.
    /// </summary>
    private static FormulaValue Npv(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double rate, out CellValue error))
            return FormulaValue.Scalar(error);
        if (rate == -1)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double total = 0;
        double discount = 1;
        foreach (ref readonly FormulaValue arg in args[1..])
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!Aggregates.TryNumeric(value, fromRange, context, out double amount, out error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                discount *= 1 + rate;
                total += amount / discount;
            }
        }

        return Finite(total);
    }

    /// <summary>
    /// <c>IRR(values, [guess])</c>: the rate at which the flows come to nothing. Unlike
    /// <c>NPV</c> the first amount sits at period zero, which is why an outlay belongs in the
    /// range rather than in front of it.
    /// </summary>
    private static FormulaValue Irr(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        var flows = new List<double>();
        foreach (CellValue value in new ArgumentValues(args[0], context))
        {
            if (Aggregates.TryNumeric(value, fromRange: true, context, out double amount, out CellValue error))
                flows.Add(amount);
            else if (error.IsError)
                return FormulaValue.Scalar(error);
        }

        if (flows.Count < 2)
            return FormulaValue.Error(ExcelError.Number);

        double guess = 0.1;
        if (args.Length > 1 && !args[1].IsBlank &&
            !Args.TryNumber(context, args[1], out guess, out CellValue failure))
        {
            return FormulaValue.Scalar(failure);
        }

        double[] amounts = [.. flows];
        return RateSolver.Solve(rate => PresentValue(amounts, rate), guess);
    }

    /// <summary>Worth of a series whose first amount falls at period zero.</summary>
    private static double PresentValue(double[] flows, double rate)
    {
        if (rate <= -1)
            return double.NaN;

        double total = 0;
        double discount = 1;
        foreach (double flow in flows)
        {
            total += flow / discount;
            discount *= 1 + rate;
        }

        return total;
    }

    /// <summary>Reads the arguments as numbers, defaulting the optional trailing ones to zero.</summary>
    private static bool ReadNumbers(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Span<double> numbers, out CellValue error)
    {
        error = CellValue.Blank;
        for (int i = 0; i < numbers.Length; i++)
        {
            if (i >= args.Length || args[i].IsBlank)
            {
                numbers[i] = 0;
                continue;
            }

            if (!Args.TryNumber(context, args[i], out double number, out error))
                return false;

            numbers[i] = number;
        }

        return true;
    }

    /// <summary>A result that ran off the end of the number line is <c>#NUM!</c>, not infinity.</summary>
    private static FormulaValue Finite(double value) =>
        double.IsNaN(value) || double.IsInfinity(value)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(value);
}
