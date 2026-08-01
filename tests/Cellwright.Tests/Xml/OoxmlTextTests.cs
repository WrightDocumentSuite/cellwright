using Cellwright.Xml;
using Xunit;

namespace Cellwright.Tests.Xml;

/// <summary>
/// The <c>_xHHHH_</c> convention that smuggles XML-invalid characters through part content.
/// Decoding has to be exact in both directions: too eager and ordinary text loses characters.
/// </summary>
public sealed class OoxmlTextTests
{
    [Theory]
    [InlineData("_x000a_", "\n")]
    [InlineData("Purchase_x000a_price", "Purchase\nprice")]
    [InlineData("_x0041_BC", "ABC")]
    [InlineData("a_x0009_b_x0009_c", "a\tb\tc")]
    [InlineData("_x005F_x000a_", "_x000a_")]
    public void EscapesAreDecoded(string encoded, string expected)
    {
        Assert.Equal(expected, OoxmlText.Decode(encoded));
    }

    [Theory]
    [InlineData("plain text")]
    [InlineData("A_B1234_C")]
    [InlineData("snake_case_name")]
    [InlineData("_x12_")]
    [InlineData("_xZZZZ_")]
    [InlineData("_x1234")]
    [InlineData("_X1234_")]
    public void WhatIsNotAnEscapeIsLeftAlone(string text)
    {
        // The same instance comes back, so an ordinary string costs nothing to decode.
        Assert.Same(text, OoxmlText.Decode(text));
    }
}
