using System.Text;
using Cellwright.Xml;
using Xunit;

namespace Cellwright.Tests.Xml;

public sealed class Utf8XmlWriterTests
{
    private static async Task<string> WriteAsync(Action<Utf8XmlWriter> write)
    {
        var stream = new MemoryStream();
        await using (var writer = new Utf8XmlWriter(stream))
        {
            write(writer);
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    [Fact]
    public async Task EscapesMarkupCharacters()
    {
        string xml = await WriteAsync(w => w.WriteText("a<b>&\"c\""));
        Assert.Equal("a&lt;b&gt;&amp;\"c\"", xml);
    }

    [Fact]
    public async Task AttributeValues_EscapeQuotesAndWhitespace()
    {
        string xml = await WriteAsync(w => w.WriteAttributeText("a\"b\tc\nd"));
        Assert.Equal("a&quot;b&#9;c&#10;d", xml);
    }

    [Fact]
    public async Task ElementText_KeepsTabsAndNewlines()
    {
        string xml = await WriteAsync(w => w.WriteText("a\tb\nc"));
        Assert.Equal("a\tb\nc", xml);
    }

    [Fact]
    public async Task InvalidControlChars_UseOoxmlEscape()
    {
        string xml = await WriteAsync(w => w.WriteText("a\u0000b\u001Fc"));
        Assert.Equal("a_x0000_b_x001F_c", xml);
    }

    [Fact]
    public async Task LiteralUnderscoreEscapeSequence_IsProtected()
    {
        string xml = await WriteAsync(w => w.WriteText("value _x0041_ stays literal"));
        Assert.Equal("value _x005F_x0041_ stays literal", xml);
    }

    [Fact]
    public async Task PlainUnderscore_PassesThrough()
    {
        string xml = await WriteAsync(w => w.WriteText("snake_case_name"));
        Assert.Equal("snake_case_name", xml);
    }

    [Fact]
    public async Task Numbers_AreInvariant()
    {
        string xml = await WriteAsync(w =>
        {
            w.WriteDouble(1234.5678);
            w.WriteRaw(" "u8);
            w.WriteInt32(-42);
        });
        Assert.Equal("1234.5678 -42", xml);
    }

    [Fact]
    public async Task Formattable_WritesCellRef()
    {
        string xml = await WriteAsync(w => w.WriteFormattable(new CellRef(12, 28)));
        Assert.Equal("AB12", xml);
    }

    [Fact]
    public async Task Unicode_TranscodesToUtf8()
    {
        string xml = await WriteAsync(w => w.WriteText("Привет 你好 🎉"));
        Assert.Equal("Привет 你好 🎉", xml);
    }

    [Fact]
    public async Task LargeContent_GrowsAndFlushes()
    {
        string chunk = new('x', 10_000);
        string xml = await WriteAsync(w =>
        {
            for (int i = 0; i < 20; i++)
                w.WriteText(chunk);
        });
        Assert.Equal(200_000, xml.Length);
    }
}
