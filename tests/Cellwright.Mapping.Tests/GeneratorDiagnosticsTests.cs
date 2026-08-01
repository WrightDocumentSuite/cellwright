using Cellwright.Mapping.Generator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Xunit;

namespace Cellwright.Mapping.Tests;

public sealed class GeneratorDiagnosticsTests
{
    private static (Diagnostic[] Diagnostics, int GeneratedSources) RunGenerator(string source)
    {
        var parseOptions = new CSharpParseOptions(LanguageVersion.Preview);
        CSharpCompilation compilation = CSharpCompilation.Create(
            "GeneratorTest",
            [CSharpSyntaxTree.ParseText(source, parseOptions)],
            [
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(ExcelRowAttribute).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(CellValue).Assembly.Location),
            ],
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        GeneratorDriver driver = CSharpGeneratorDriver.Create(
            [new RowSchemaGenerator().AsSourceGenerator()],
            parseOptions: parseOptions);
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out _, out _);
        GeneratorDriverRunResult result = driver.GetRunResult();
        return ([.. result.Diagnostics], result.GeneratedTrees.Length);
    }

    [Fact]
    public void NonPartialType_ReportsCwm001()
    {
        (Diagnostic[] diagnostics, int generated) = RunGenerator(
            """
            using Cellwright.Mapping;

            [ExcelRow]
            public record NotPartial(int Id);
            """);

        Assert.Contains(diagnostics, d => d.Id == "CWM001");
        Assert.Equal(0, generated);
    }

    [Fact]
    public void UnsupportedPropertyType_ReportsCwm007()
    {
        (Diagnostic[] diagnostics, _) = RunGenerator(
            """
            using Cellwright.Mapping;

            [ExcelRow]
            public partial record Bad(int Id, System.Uri Address);
            """);

        Assert.Contains(diagnostics, d => d.Id == "CWM007");
    }

    [Fact]
    public void ValidRecord_GeneratesOneSchemaWithoutDiagnostics()
    {
        (Diagnostic[] diagnostics, int generated) = RunGenerator(
            """
            using Cellwright.Mapping;

            namespace Testing;

            [ExcelRow]
            public partial record Good(int Id, string Name, double? Score);
            """);

        Assert.Empty(diagnostics);
        Assert.Equal(1, generated);
    }
}
