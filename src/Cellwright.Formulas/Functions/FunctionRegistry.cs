using System.Collections.Frozen;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// Collects function definitions while the registry is built.
/// </summary>
internal sealed class FunctionRegistryBuilder
{
    private readonly Dictionary<string, FunctionDefinition> _functions = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Registers a function; the last registration of a name wins.</summary>
    public FunctionRegistryBuilder Add(
        string name, int minArgs, int maxArgs, FunctionHandler handler,
        bool isVolatile = false, bool wantsRawReferences = false, bool acceptsErrors = false,
        bool dependsOnVisibility = false)
    {
        _functions[name] = new FunctionDefinition(name, minArgs, maxArgs, handler)
        {
            IsVolatile = isVolatile,
            WantsRawReferences = wantsRawReferences,
            AcceptsErrors = acceptsErrors,
            DependsOnVisibility = dependsOnVisibility,
        };
        return this;
    }

    /// <summary>
    /// Registers another name for a function that is already in — the renamed statistics of
    /// 2010 (<c>STDEV.S</c> for <c>STDEV</c>) and the like. The alias must come after its
    /// target's registration.
    /// </summary>
    public FunctionRegistryBuilder Alias(string name, string existing)
    {
        FunctionDefinition target = _functions[existing];
        _functions[name] = new FunctionDefinition(name, target.MinArgs, target.MaxArgs, target.Handler)
        {
            IsVolatile = target.IsVolatile,
            WantsRawReferences = target.WantsRawReferences,
            AcceptsErrors = target.AcceptsErrors,
            DependsOnVisibility = target.DependsOnVisibility,
        };
        return this;
    }

    internal FrozenDictionary<string, FunctionDefinition> Build() =>
        _functions.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
}

/// <summary>
/// All worksheet functions the engine knows, resolved by name without allocating: lookups
/// run straight off the lexer's span.
/// </summary>
internal static class FunctionRegistry
{
    private static readonly FrozenDictionary<string, FunctionDefinition> Functions = Build();

    private static readonly FrozenDictionary<string, FunctionDefinition>.AlternateLookup<ReadOnlySpan<char>> SpanLookup =
        Functions.GetAlternateLookup<ReadOnlySpan<char>>();

    /// <summary>Names of every registered function, for documentation and coverage reports.</summary>
    public static IReadOnlyCollection<string> Names => Functions.Keys;

    /// <summary>
    /// Resolves a function by name, ignoring the compatibility prefixes a file spells post-2007
    /// functions with. Formulas normally arrive bare — the reader strips them — but a formula
    /// handed straight to the parser may still carry one.
    /// </summary>
    public static FunctionDefinition? Resolve(ReadOnlySpan<char> name)
    {
        if (name.StartsWith(FutureFunctions.Prefix, StringComparison.OrdinalIgnoreCase))
        {
            name = name[FutureFunctions.Prefix.Length..];
            if (name.StartsWith(FutureFunctions.WorksheetPrefix, StringComparison.OrdinalIgnoreCase))
                name = name[FutureFunctions.WorksheetPrefix.Length..];
        }

        return SpanLookup.TryGetValue(name, out FunctionDefinition? definition) ? definition : null;
    }

    private static FrozenDictionary<string, FunctionDefinition> Build()
    {
        var builder = new FunctionRegistryBuilder();
        MathFunctions.Register(builder);
        ExtendedMathFunctions.Register(builder);
        MatrixFunctions.Register(builder);
        LogicalFunctions.Register(builder);
        TextFunctions.Register(builder);
        NumberTextFunctions.Register(builder);
        DateTimeFunctions.Register(builder);
        DayCountFunctions.Register(builder);
        InformationFunctions.Register(builder);
        LookupFunctions.Register(builder);
        VectorLookupFunctions.Register(builder);
        ExtendedStatisticalFunctions.Register(builder);
        CorrelationFunctions.Register(builder);
        TrendFunctions.Register(builder);
        EngineeringFunctions.Register(builder);
        ComplexFunctions.Register(builder);
        BesselFunctions.Register(builder);
        AnnualRateFunctions.Register(builder);
        CouponFunctions.Register(builder);
        SecurityFunctions.Register(builder);
        OddPeriodFunctions.Register(builder);
        ConnectionStubFunctions.Register(builder);
        CellInformationFunctions.Register(builder);
        ConditionalFunctions.Register(builder);
        StatisticalFunctions.Register(builder);
        AggregateFunctions.Register(builder);
        FinancialFunctions.Register(builder);
        DatabaseFunctions.Register(builder);
        RegressionFunctions.Register(builder);
        DynamicArrayFunctions.Register(builder);
        ArrayBuilderFunctions.Register(builder);
        DistributionFunctions.Register(builder);
        CashFlowFunctions.Register(builder);
        DepreciationFunctions.Register(builder);
        RegisterModernStatisticsNames(builder);
        return builder.Build();
    }

    /// <summary>
    /// The 2010 renames of the statistics catalog that are pure renames. The ones whose new
    /// name changed behaviour (<c>CHISQ.DIST</c> against <c>CHIDIST</c>) have their own
    /// registrations instead.
    /// </summary>
    private static void RegisterModernStatisticsNames(FunctionRegistryBuilder builder) => builder
        .Alias("STDEV.S", "STDEV")
        .Alias("STDEV.P", "STDEVP")
        .Alias("VAR.S", "VAR")
        .Alias("VAR.P", "VARP")
        .Alias("MODE.SNGL", "MODE")
        .Alias("RANK.EQ", "RANK")
        // The name ISO/IEC 29500 gave the function Excel spells ISO.CEILING.
        .Alias("ECMA.CEILING", "ISO.CEILING")
        // The byte-counting twins measure DBCS text in bytes; in single-byte text — which is
        // what a culture-independent engine works in — they equal their character twins.
        .Alias("LEFTB", "LEFT")
        .Alias("RIGHTB", "RIGHT")
        .Alias("MIDB", "MID")
        .Alias("LENB", "LEN")
        .Alias("FINDB", "FIND")
        .Alias("SEARCHB", "SEARCH")
        .Alias("REPLACEB", "REPLACE");
}
