using Cellwright.Formulas.Functions;

namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Precedence-climbing parser producing a reverse-Polish program. Precedence follows Excel,
/// including its quirk that unary minus binds tighter than exponentiation, so
/// <c>-2^2</c> is 4.
/// </summary>
internal ref struct FormulaParser
{
    /// <summary>Nesting limit. Excel stops at 64; the extra headroom still rules out stack overflow.</summary>
    private const int MaxDepth = 128;

    private const int UnaryPrecedence = 6;

    private readonly FormulaProgramBuilder _builder;
    private readonly string _text;
    private readonly Workbook? _workbook;
    private readonly List<(string Name, int Slot)> _bindings;
    private FormulaLexer _lexer;
    private FormulaToken _current;
    private int _depth;

    private FormulaParser(string text, Workbook? workbook)
    {
        _text = text;
        _workbook = workbook;
        _builder = new FormulaProgramBuilder();
        _bindings = [];
        _lexer = new FormulaLexer(text);
        _current = _lexer.Next();
    }

    /// <summary>
    /// Compiles formula text. Sheet names resolve against <paramref name="workbook"/>;
    /// unknown sheets survive as unresolved references that evaluate to <c>#REF!</c>.
    /// </summary>
    public static FormulaProgram Parse(string text, Workbook? workbook = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        string body = text.StartsWith('=') ? text[1..] : text;
        var parser = new FormulaParser(body, workbook);
        return parser.ParseProgram(body);
    }

    private FormulaProgram ParseProgram(string body)
    {
        if (_current.Kind == TokenKind.End)
            throw new FormulaParseException("The formula is empty", _text, 0);

        ParseExpression(1);
        if (_current.Kind != TokenKind.End)
            throw new FormulaParseException($"Unexpected '{_current.Text}'", _text, _current.Start);

        return _builder.Build(body);
    }

    private void ParseExpression(int minPrecedence)
    {
        EnterDepth();
        ParseUnary();
        while (true)
        {
            int precedence = BinaryPrecedence(_current.Kind);
            if (precedence == 0 || precedence < minPrecedence)
                break;

            TokenKind op = _current.Kind;
            Advance();
            ParseExpression(precedence + 1);
            _builder.Emit(BinaryOpCode(op));
        }

        _depth--;
    }

    private void ParseUnary()
    {
        if (_current.Kind is TokenKind.Minus or TokenKind.Plus)
        {
            TokenKind op = _current.Kind;
            Advance();
            ParseExpression(UnaryPrecedence);
            _builder.Emit(op == TokenKind.Minus ? OpCode.Negate : OpCode.Plus);
            return;
        }

        // A leading @ asks for the one value of what follows. It is written as an operator but
        // it is the SINGLE function, and Excel spells it that way too in older files.
        if (_current.Kind == TokenKind.At)
        {
            Advance();
            ParseExpression(UnaryPrecedence);
            _builder.EmitCall(FunctionRegistry.Resolve("SINGLE"), "SINGLE", 1);
            return;
        }

        ParsePrimary();
        while (_current.Kind is TokenKind.Percent or TokenKind.OpenParen)
        {
            if (_current.Kind == TokenKind.Percent)
            {
                Advance();
                _builder.Emit(OpCode.Percent);
            }
            else
            {
                Advance();
                int argCount = ParseArguments();
                Expect(TokenKind.CloseParen, ")");
                _builder.EmitInvoke(argCount);
            }
        }
    }

    private void ParsePrimary()
    {
        switch (_current.Kind)
        {
            case TokenKind.Number:
                _builder.EmitNumber(double.Parse(_current.Text, System.Globalization.CultureInfo.InvariantCulture));
                Advance();
                return;

            case TokenKind.String:
                _builder.EmitString(_current.Text.ToString());
                Advance();
                return;

            case TokenKind.Boolean:
                _builder.EmitBoolean(_current.Text.Equals("TRUE", StringComparison.OrdinalIgnoreCase));
                Advance();
                return;

            case TokenKind.Error:
                ExcelErrorText.TryParse(_current.Text, out ExcelError error);
                _builder.EmitError(error);
                Advance();
                return;

            case TokenKind.Reference:
                EmitReferenceToken();
                Advance();
                return;

            case TokenKind.Name:
                ParseNameOrCall();
                return;

            case TokenKind.TableReference:
                EmitTableReference(tableName: null, _current.Start);
                Advance();
                return;

            case TokenKind.OpenParen:
                ParseParenthesized();
                return;

            case TokenKind.OpenBrace:
                ParseArrayLiteral();
                return;

            default:
                throw new FormulaParseException($"Expected a value but found '{_current.Text}'", _text, _current.Start);
        }
    }

    private void ParseParenthesized()
    {
        Advance();
        ParseExpression(1);

        // A comma between parentheses joins reference areas; inside a call it separates arguments,
        // which is why unions are only recognized here.
        while (_current.Kind == TokenKind.Comma)
        {
            Advance();
            ParseExpression(1);
            _builder.Emit(OpCode.Union);
        }

        Expect(TokenKind.CloseParen, ")");
    }

    private void ParseNameOrCall()
    {
        string name = _current.Text.ToString();
        int nameStart = _current.Start;
        Advance();
        if (_current.Kind == TokenKind.TableReference)
        {
            EmitTableReference(name, nameStart);
            Advance();
            return;
        }

        if (_current.Kind != TokenKind.OpenParen)
        {
            if (TryEmitLocal(name))
                return;

            _builder.EmitName(name);
            return;
        }

        if (IsLet(name))
        {
            ParseLet();
            return;
        }

        if (IsLambda(name))
        {
            ParseLambda();
            return;
        }

        if (TryEmitLocal(name))
            return;

        Advance();
        int argCount = ParseArguments();
        Expect(TokenKind.CloseParen, ")");

        FunctionDefinition? definition = FunctionRegistry.Resolve(name);
        if (definition is not null && (argCount < definition.MinArgs || argCount > definition.MaxArgs))
        {
            throw new FormulaParseException(
                $"{definition.Name} takes {DescribeArity(definition)} but got {argCount}", _text, nameStart);
        }

        _builder.EmitCall(definition, name, argCount);
    }

    /// <summary>
    /// Parses <c>LET(name, value, …, calculation)</c>. The names are not expressions, so this
    /// cannot go through the ordinary argument loop: each pair becomes a slot the calculation —
    /// and every later value — can read, and the whole thing is one operand on the stack.
    /// </summary>
    private void ParseLet()
    {
        int start = _current.Start;
        Advance();
        int scopeStart = _bindings.Count;
        while (_current.Kind == TokenKind.Name && NextTokenIsComma())
        {
            string binding = StripLocalPrefix(_current.Text.ToString());
            Advance();
            Advance();
            ParseExpression(1);

            // Declared after its value is parsed, so a binding cannot refer to itself.
            int slot = _builder.DeclareLocal(binding);
            _builder.Emit(OpCode.StoreLocal, slot);
            _bindings.Add((binding, slot));
            Expect(TokenKind.Comma, ",");
        }

        int count = _bindings.Count - scopeStart;
        if (count == 0)
            throw new FormulaParseException("LET needs at least one name and value before the calculation", _text, start);

        ParseExpression(1);
        Expect(TokenKind.CloseParen, ")");
        _bindings.RemoveRange(scopeStart, count);
        _builder.EmitLet(count);
    }

    private void ParseLambda()
    {
        int start = _current.Start;
        Advance();
        if (_current.Kind == TokenKind.CloseParen)
            throw new FormulaParseException("LAMBDA needs a calculation expression", _text, start);

        int scopeStart = _bindings.Count;
        var parameterSlots = new List<int>();
        while (_current.Kind == TokenKind.Name && NextTokenIsComma())
        {
            string paramName = StripLocalPrefix(_current.Text.ToString());
            Advance();
            Advance();
            int slot = _builder.DeclareLocal(paramName);
            parameterSlots.Add(slot);
            _bindings.Add((paramName, slot));
        }

        int lambdaIndex = _builder.BeginLambda([.. parameterSlots]);
        ParseExpression(1);
        Expect(TokenKind.CloseParen, ")");
        _bindings.RemoveRange(scopeStart, parameterSlots.Count);
        _builder.EndLambda(lambdaIndex);
    }

    /// <summary>Emits a read of a <c>LET</c> or <c>LAMBDA</c> binding, innermost first. Returns whether the name was bound.</summary>
    private bool TryEmitLocal(string name)
    {
        string normalized = StripLocalPrefix(name);
        for (int i = _bindings.Count - 1; i >= 0; i--)
        {
            if (!string.Equals(_bindings[i].Name, normalized, StringComparison.OrdinalIgnoreCase))
                continue;

            _builder.Emit(OpCode.PushLocal, _bindings[i].Slot);
            return true;
        }

        return false;
    }

    private static string StripLocalPrefix(string name) =>
        name.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name;

    /// <summary>
    /// Whether the token after the current one is a comma, which is what tells a binding name
    /// apart from a calculation that happens to start with a name.
    /// </summary>
    private bool NextTokenIsComma()
    {
        FormulaLexer saved = _lexer;
        bool comma = _lexer.Next().Kind == TokenKind.Comma;
        _lexer = saved;
        return comma;
    }

    private static bool IsLet(string name) =>
        name.Equals("LET", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("_xlfn.LET", StringComparison.OrdinalIgnoreCase);

    private static bool IsLambda(string name) =>
        name.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("_xlfn.LAMBDA", StringComparison.OrdinalIgnoreCase);

    private int ParseArguments()
    {
        if (_current.Kind == TokenKind.CloseParen)
            return 0;

        int count = 0;
        while (true)
        {
            if (_current.Kind is TokenKind.Comma or TokenKind.CloseParen)
                _builder.Emit(OpCode.PushBlank); // Omitted argument, as in IF(a,,b).
            else
                ParseExpression(1);

            count++;
            if (_current.Kind != TokenKind.Comma)
                return count;
            Advance();
        }
    }

    private void ParseArrayLiteral()
    {
        Advance();
        var values = new List<CellValue>();
        int columns = 0;
        int rows = 1;
        int rowLength = 0;

        while (true)
        {
            values.Add(ParseArrayElement());
            rowLength++;
            if (_current.Kind == TokenKind.Comma)
            {
                Advance();
                continue;
            }

            columns = Math.Max(columns, rowLength);
            if (_current.Kind != TokenKind.Semicolon)
                break;

            Advance();
            rows++;
            rowLength = 0;
        }

        Expect(TokenKind.CloseBrace, "}");
        if (values.Count != rows * columns)
            throw new FormulaParseException("Array constant rows must all have the same length", _text, _current.Start);

        _builder.EmitArray(new ArrayLiteral(rows, columns, [.. values]));
    }

    private CellValue ParseArrayElement()
    {
        bool negate = false;
        if (_current.Kind is TokenKind.Minus or TokenKind.Plus)
        {
            negate = _current.Kind == TokenKind.Minus;
            Advance();
        }

        CellValue value = _current.Kind switch
        {
            TokenKind.Number => CellValue.Number(double.Parse(_current.Text, System.Globalization.CultureInfo.InvariantCulture)),
            TokenKind.String => CellValue.Text(_current.Text.ToString()),
            TokenKind.Boolean => CellValue.Boolean(_current.Text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)),
            TokenKind.Error => ParseErrorLiteral(),
            _ => throw new FormulaParseException("Array constants may only contain literals", _text, _current.Start),
        };

        Advance();
        return negate && value.IsNumber ? CellValue.Number(-value.GetNumber()) : value;
    }

    private CellValue ParseErrorLiteral()
    {
        ExcelErrorText.TryParse(_current.Text, out ExcelError error);
        return CellValue.Error(error);
    }

    private void EmitReferenceToken()
    {
        string? sheetName = _lexer.SheetName;
        string? spanEnd = _lexer.SheetSpanEnd;
        int externalBook = _lexer.ExternalBook;
        RefPart start = _lexer.Start;
        RefPart end = _lexer.End;

        int sheetId = FormulaRef.CurrentSheet;
        int sheetNameIndex = -1;
        if (sheetName is not null)
        {
            sheetNameIndex = _builder.InternString(sheetName);

            // A sheet of another workbook never resolves against this one's tabs; the book
            // index and the name travel to the evaluator, which reads the cached values.
            Worksheet? target = externalBook > 0 ? null : _workbook?.Sheets.Find(sheetName);
            sheetId = target?.Id ?? FormulaRef.UnresolvedSheet;
        }

        int spanSheetId = FormulaRef.NoSheet;
        int spanSheetNameIndex = -1;
        if (spanEnd is not null)
        {
            spanSheetNameIndex = _builder.InternString(spanEnd);
            Worksheet? target = externalBook > 0 ? null : _workbook?.Sheets.Find(spanEnd);
            spanSheetId = target?.Id ?? FormulaRef.UnresolvedSheet;
        }

        _builder.EmitReference(new FormulaRef(
            sheetId, sheetNameIndex, start, end, spanSheetId, spanSheetNameIndex, externalBook,
            externalBook > 0 ? sheetName : null, externalBook > 0 ? spanEnd : null,
            IsSpilled: _lexer.SpilledReference));
    }

    /// <summary>
    /// Emits the structured reference the lexer just produced. The whole reference, table name
    /// included, is kept as text so that rendering the program back gives the original spelling.
    /// </summary>
    private void EmitTableReference(string? tableName, int start)
    {
        int end = _current.Start + _current.Length;
        string text = _text[start..end];
        _builder.EmitTableReference(TableRef.Parse(tableName, _current.Text, text));
    }

    private void EnterDepth()
    {
        if (++_depth > MaxDepth)
            throw new FormulaParseException($"The formula nests deeper than {MaxDepth} levels", _text, _current.Start);
    }

    private void Advance() => _current = _lexer.Next();

    private void Expect(TokenKind kind, string expected)
    {
        if (_current.Kind != kind)
            throw new FormulaParseException($"Expected '{expected}'", _text, _current.Start);
        Advance();
    }

    private static string DescribeArity(FunctionDefinition definition) =>
        definition.MaxArgs == FunctionDefinition.Unlimited
            ? $"at least {definition.MinArgs} argument(s)"
            : definition.MinArgs == definition.MaxArgs
                ? $"{definition.MinArgs} argument(s)"
                : $"{definition.MinArgs}-{definition.MaxArgs} arguments";

    private static int BinaryPrecedence(TokenKind kind) => kind switch
    {
        TokenKind.Equal or TokenKind.NotEqual or TokenKind.Less or
        TokenKind.LessOrEqual or TokenKind.Greater or TokenKind.GreaterOrEqual => 1,
        TokenKind.Ampersand => 2,
        TokenKind.Plus or TokenKind.Minus => 3,
        TokenKind.Star or TokenKind.Slash => 4,
        TokenKind.Caret => 5,
        TokenKind.Colon or TokenKind.Space => 8,
        _ => 0,
    };

    private static OpCode BinaryOpCode(TokenKind kind) => kind switch
    {
        TokenKind.Plus => OpCode.Add,
        TokenKind.Minus => OpCode.Subtract,
        TokenKind.Star => OpCode.Multiply,
        TokenKind.Slash => OpCode.Divide,
        TokenKind.Caret => OpCode.Power,
        TokenKind.Ampersand => OpCode.Concat,
        TokenKind.Equal => OpCode.Equal,
        TokenKind.NotEqual => OpCode.NotEqual,
        TokenKind.Less => OpCode.Less,
        TokenKind.LessOrEqual => OpCode.LessOrEqual,
        TokenKind.Greater => OpCode.Greater,
        TokenKind.GreaterOrEqual => OpCode.GreaterOrEqual,
        TokenKind.Colon => OpCode.Range,
        TokenKind.Space => OpCode.Intersect,
        _ => throw new InvalidOperationException($"{kind} is not a binary operator."),
    };
}
