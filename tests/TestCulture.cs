using System.Globalization;
using System.Runtime.CompilerServices;

namespace Cellwright.Tests.Shared;

/// <summary>
/// Runs a test assembly under whatever culture <c>CELLWRIGHT_TEST_CULTURE</c> names.
/// </summary>
/// <remarks>
/// A spreadsheet file is culture-neutral: every number in it is written with a full stop and
/// read back the same way, whoever opens it. That only holds while every parse and every format
/// says so, and one call that forgets produces a file that is wrong in Berlin and right in
/// Boston. Tests under a single culture cannot see the difference, so the build runs them twice:
/// once as they are, once under a culture whose decimal separator is a comma.
/// </remarks>
internal static class TestCulture
{
    [ModuleInitializer]
    internal static void Apply()
    {
        if (Environment.GetEnvironmentVariable("CELLWRIGHT_TEST_CULTURE") is not { Length: > 0 } name)
            return;

        var culture = CultureInfo.GetCultureInfo(name);
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
    }
}
