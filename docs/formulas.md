# Formulas (Cellwright.Formulas)

The `Cellwright.Formulas` package adds formula evaluation to the document model: a parser,
an executor, and incremental recalculation over the dependency graph. The core knows nothing
about the engine — referencing the package does not change behavior for callers that do not
use it.

## Quick start

```csharp
using Cellwright;
using Cellwright.Formulas;

var book = Workbook.Create();
Worksheet sheet = book.Sheets.Add("Report");
sheet["A1"] = 10;
sheet["A2"] = 20;
sheet.Cell("A3").SetFormula("SUM(A1:A2)*1.2");

CalculationSummary summary = book.Calculate();
Console.WriteLine(sheet["A3"].GetNumber());  // 36
Console.WriteLine(summary);                  // 1 cell(s) in 0.3 ms, 0 error(s), 0 circular
```

Calculation is explicit: you call `Calculate()`, as in EPPlus and ClosedXML. The first call
computes everything; later calls recompute only what changed, plus volatile formulas and
everything that depends on them.

```csharp
sheet["A1"] = 100;
book.Calculate();                                   // only A3 is recalculated
book.Calculate(new CalculateOptions { Force = true }); // force a full recalculation
```

Additionally:

```csharp
// One-off evaluation of an expression that is not stored anywhere
CellValue total = book.Evaluate("SUM(Report!A1:A2)*1.2");
CellValue local = sheet.Evaluate("A1+1", originCell: "B5"); // relative refs from B5
CellValue copied = sheet.EvaluateCopied(                 // formula copied B1 -> B500
    "$A1>0", new CellRef(1, 2), new CellRef(500, 2));

// Value with lazy recalculation
double value = sheet.Cell("A3").CalculatedValue.GetNumber();

// Diagnostics and cancellation
book.Calculate(new CalculateOptions
{
    OnWarning = w => log.Warn(w),      // cycles, broken formulas, refs to deleted sheets
    CancellationToken = token,
});
```

`Calculate` is synchronous: the work is entirely CPU-bound, and wrapping it in `Task` would
be misleading. `Workbook` is not thread-safe; the engine is single-threaded.

`EvaluateCopied` keeps the original formula text as the compiled-program cache key and applies the
row/column copy offset only while executing it. This is the intended callback for conditional
formatting over a large range: mixed references retain their `$` axes without creating one
translated formula string and one cached program per target. Each public call is an independent
calculation pass; volatile `NOW()`/`TODAY()` values are therefore frozen within a call, not across a
caller-managed loop of calls.

## Table references and local names

```csharp
sheet.Tables.Add("Sales", "A1:C100");

sheet.Cell("E1").SetFormula("SUM(Sales[Amount])");            // column body
sheet.Cell("E2").SetFormula("SUM(Sales[[#Totals],[Amount]])"); // totals row
sheet.Cell("E3").SetFormula("SUM(Sales[[Qty]:[Amount]])");  // several adjacent columns
sheet.Cell("D5").SetFormula("[@Qty]*[@Price]");               // current row, inside the table

// LET evaluates an expression once and gives it a name
sheet.Cell("E4").SetFormula("LET(total,SUM(Sales[Amount]),IF(total>0,total,\"no data\"))");

// TEXT renders a value with a number format
sheet.Cell("E5").SetFormula("\"Total: \"&TEXT(E1,\"#,##0.00\")");
```

A structured reference survives inserting a row in the table without changing: it names the
table, not coordinates, and resolves to its current shape on every recalculation.

## 3-D ranges and other workbooks

```csharp
sheet.Cell("B1").SetFormula("SUM(Q1:Q4!A1)");        // one rectangle across four sheets
sheet.Cell("B2").SetFormula("SUM('S 1:S 3'!A1:C9)"); // names with spaces — in single quotes
sheet.Cell("B3").SetFormula("[1]Price!B2*C3");       // closed workbook — from link cache
```

A 3-D range (§18.17.2.3) reads one rectangle from every sheet between the endpoints in tab
order, including endpoints written in reverse order; each sheet is subscribed separately in the
dependency graph, so editing a middle sheet recalculates the total. Renaming an endpoint sheet
updates the formula text; inserting rows does not move a 3-D range — one shift cannot be correct
for all sheets with their different geometry.

A reference to another workbook (`[1]Sheet1!A1`) answers from values the file cached in
`externalLink` parts (§18.14) at the last link refresh — exactly as Excel does while the source
is closed. The foreign file itself is never opened. The cache is also available from code:
`book.ExternalBooks[0].GetCachedValue("Sheet1", 1, 1)`. A workbook without a cache yields `#REF!`.

## What happens to the file

- **Shared formulas** (`<f t="shared" si="…"/>`) are preserved as a group on load, and the first
  `Calculate()` expands them into ordinary formulas with shifted relative references. Without
  `Cellwright.Formulas`, group members remain without formula text — the core warns about this.
- **Array formulas** keep the `ref` attribute and are written back as `t="array"`. Dynamic array
  formulas spill into neighbouring blank cells (`Worksheet.GetSpillRange` / `GetSpillOrigin`), or
  yield `#SPILL!` when blocked.
- **`fullCalcOnLoad`**: after a successful recalculation with no cycles, the flag is not written,
  and Excel opens the file with our values. An uncalculated workbook keeps the flag.
- **`calcChain.xml`** is ignored on read and not created on write: our own graph is more reliable,
  and Excel rebuilds the chain itself.
- **`_xlfn.` prefix**: a function absent from ISO/IEC 29500 must be written as
  `_xlfn.XLOOKUP(...)` — bare `XLOOKUP(` yields `#NAME?` in Excel even though Excel knows the
  function (MS-XLSX 2.2.3). This is a format-writing quirk, not a formula quirk, so read strips
  the prefix and write restores it: `cell.Formula` always shows the name a person would type,
  and `SetFormula("XLOOKUP(...)")` produces a file Excel understands. Three functions —
  `FILTER`, `SORT`, `PY` — are sheet-only and carry a second prefix, `_xlfn._xlws.`. A prefixed
  name not in the table is left alone entirely: it comes from a newer Excel version, and stripping
  its prefix would break the file.

## Cycles, errors, volatile

- A circular reference does not throw: cycle participants get 0, addresses land in
  `CalculationSummary.CircularCells`, and an `OnWarning` `CircularReference` arrives for each
  cell.
- An unparseable formula yields `#NAME?` and a `BadFormula` warning; a reference to a
  nonexistent sheet yields `#REF!`.
- Volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`) recalculate
  every pass. `NOW()` and `TODAY()` are identical within one pass.

## Excel semantics implemented deliberately

These rules are easy to miss, so each has its own test group.

Separately, there is a family of rules where Excel diverges from ISO/IEC 29500: `POWER(0,0)`
yields `#NUM!`, not `#DIV/0!`; `CEILING(-4.2,2)` refuses to compute rather than rounding toward
zero; `LOG(10,1)` yields `#DIV/0!`; `MATCH` with no match yields `#N/A`, not `#NUM!`. The engine
follows Excel. The full list of eighteen divergences, with MS-OI29500 section references, is in
[conformance.md](conformance.md).

- **Empty cell**: 0 in arithmetic, `""` in concatenation; for an empty cell both `A1=0` and
  `A1=""` yield TRUE, but `ISBLANK` distinguishes them. A formula that returns a reference to an
  empty cell yields 0 — an empty result in a formula cell is impossible.
- **Type order**: number < text < FALSE < TRUE. Text compares case-insensitively; `EXACT` is
  case-sensitive.
- **Rounding**: `ROUND` rounds halves away from zero (not banker's rounding) and works on the
  15-digit representation of the number, so `ROUND(1.005,2)` = 1.01, as in Excel.
- **`MOD`** takes the sign of the divisor: `MOD(-3,2)` = 1.
- **Aggregates**: text and logical values inside a range are skipped, but those written directly
  in the formula are coerced. `SUM(TRUE)` = 1; `SUM(range with TRUE)` ignores it.
- **`SUBTOTAL`/`AGGREGATE`** ignore cells that themselves contain SUBTOTAL or AGGREGATE — otherwise
  a total would double-count intermediate totals. Codes 101–111 and options 1/3/5/7 additionally
  skip hidden rows; changing row visibility marks such formulas for recalculation even though no
  cell changed.
- **Filter-hidden rows** are skipped by SUBTOTAL for any code, not only 101–111: that is the point
  of `SUBTOTAL(9,…)` under AutoFilter. In the file such rows are marked with ordinary `hidden`, so
  the signal is the presence of criteria (`filterColumn`) on the sheet AutoFilter or table — then
  hidden rows inside their range count as filtered (`Worksheet.AutoFilterIsActive`). Moving
  AutoFilter to another range clears the signal. AGGREGATE does not make this distinction: for it,
  hidden rows are skipped only by options 1/3/5/7.
- **Code range**: SUBTOTAL accepts 1–11 and 101–111; anything else is `#VALUE!`. Codes 12–19
  (MEDIAN, MODE, LARGE, SMALL, PERCENTILE.INC, QUARTILE.INC, PERCENTILE.EXC, QUARTILE.EXC) exist
  only in AGGREGATE.
- **Whole-column references** `A:A` are trimmed to the sheet used range: without that, one
  `SUM(A:A)` would walk a million rows.
- **Structured references** `Table[Column]` resolve to the table's current shape at evaluation time,
  not to coordinates at parse time: a table that grew by one row moves every formula that reads it,
  while the formulas themselves stay unchanged. Supported areas: `#All`, `#Data`, `#Headers`,
  `#Totals`, `#This Row`, the `[@Column]` form, column ranges `[[First]:[Last]]`, and apostrophes
  escaping brackets in names. Column names compare character by character, including spaces: column
  `"Amount "` exists, and `Table[Amount]` will not match it — as in Excel. Unknown table or column
  yields `#REF!`.
- **`LET`** binds names for one formula: `LET(x,A1*2,y,x+1,x*y)`. Each value sees names declared
  before it, but not itself; nested `LET` shadows outer names; outside its call the name is an
  ordinary defined name again.
- **`LAMBDA` and higher-order functions**: `LAMBDA([param1, ...], body)` creates a first-class
  closure that captures outer `LET`/`LAMBDA` lexical bindings, can be stored in a workbook or
  sheet Defined Name (`book.DefinedNames.Add("AddTax", "LAMBDA(x, rate, x*(1+rate))")`), invoked
  directly (`LAMBDA(x, x+1)(5)`), or passed to `MAP`, `REDUCE`, `SCAN`, `MAKEARRAY`, `BYROW`, and
  `BYCOL`. Omitted arguments in a lambda call are detected via `ISOMITTED(param)`. On save/load,
  `_xlfn.` and `_xlpm.` parameter prefixes are stripped on read and restored on write.
- **Implicit intersection**: a range in scalar context narrows to the formula row or column; empty
  intersection yields `#VALUE!`.
- **Criteria** for `SUMIF`/`COUNTIF` understand `">10"`, `"<>"`, `"apple*"`, and escaping with `~*`.
- **Approximate lookup does not cross type boundaries.** `VLOOKUP`, `HLOOKUP`, and `MATCH` without
  an exact match take the nearest value not above the key, but search only among values of the same
  type. Text sorts above any number, so `VLOOKUP("7",A1:A4,1)` over a column of numbers answers
  `#N/A`, not the last number in the column.
- **An error at the end of a range propagates outward**: `SUM(#REF!:B4)` is `#REF!`, not a type
  mismatch. A range whose corner was deleted should say so.
- **`AND` and `OR` read range contents, not its size**: empty cells and text inside are skipped, so
  a range padded with empties answers the same as the values in it. A range with nothing to read at
  all yields `#VALUE!`.
- **Text longer than 32767 characters** becomes `#VALUE!`.
- Formula parsing is culture-independent: in xlsx the argument separator is always `,`, and the
  decimal separator is always `.`.

## Supported functions (468)

Of the 356 functions in §18.17.7 ISO/IEC 29500, 352 are implemented — all except four tied to
Japanese and Thai locales (`ASC`, `JIS`, `PHONETIC`, `BAHTTEXT`). Beyond the standard, modern
names and dynamic-array functions are accepted.

- **Math**: SUM, PRODUCT, ABS, SIGN, SQRT, SQRTPI, POWER, EXP, LN, LOG, LOG10, MOD, QUOTIENT,
  INT, TRUNC, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, CEILING.MATH, FLOOR.MATH, CEILING.PRECISE,
  FLOOR.PRECISE, ISO.CEILING (=ECMA.CEILING), MROUND, PI, RAND, RANDBETWEEN, trigonometry with
  hyperbolic and inverse branches (SIN…ACOTH, SEC/CSC/COT), DEGREES, RADIANS, EVEN, ODD, FACT,
  FACTDOUBLE, COMBIN, PERMUT, MULTINOMIAL, SERIESSUM, GCD, LCM, SUMSQ, SUMX2MY2, SUMX2PY2,
  SUMXMY2.
- **Matrix**: MMULT, MINVERSE, MDETERM.
- **Logic & Lambda**: IF, IFS, IFERROR, IFNA, AND, OR, XOR, NOT, TRUE, FALSE, SWITCH, LET, LAMBDA,
  ISOMITTED, MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL.
- **Text**: CONCAT, CONCATENATE, TEXTJOIN, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, LEN, LEFT, RIGHT,
  MID, TRIM, CLEAN, UPPER, LOWER, PROPER, SUBSTITUTE, REPLACE, FIND, SEARCH, REPT, EXACT, CHAR,
  CODE, VALUE, T, TEXT, FIXED, DOLLAR, USDOLLAR, ROMAN (all five forms), ARABIC. Byte doubles
  (`LEFTB` and kin) are synonyms of character-wise functions: in single-byte text they match
  Excel too.
- **Date and time**: TODAY, NOW, DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY,
  WEEKNUM, EDATE, EOMONTH, DAYS, DAYS360, DATEDIF, DATEVALUE, TIMEVALUE, WORKDAY, NETWORKDAYS,
  WORKDAY.INTL, NETWORKDAYS.INTL, YEARFRAC (all five bases).
- **Information**: ISBLANK, ISNUMBER, ISTEXT, ISNONTEXT, ISLOGICAL, ISERROR, ISERR, ISNA, ISEVEN,
  ISODD, ISREF, ISFORMULA, FORMULATEXT, N, NA, TYPE, ERROR.TYPE, CELL (address, col, row,
  contents, type, format, protect, width, prefix, color, parentheses), INFO (numfile, recalc,
  system; machine answers are honest `#N/A`).
- **Lookup**: VLOOKUP, HLOOKUP, LOOKUP (vector and array forms), INDEX, MATCH, XLOOKUP, XMATCH,
  CHOOSE, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS, AREAS, ADDRESS, HYPERLINK.
- **Conditional aggregates**: SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS,
  MINIFS, COUNTBLANK, SUMPRODUCT.
- **Statistics**: COUNT, COUNTA, AVERAGE, AVERAGEA, MIN, MINA, MAX, MAXA, MEDIAN, MODE
  (=MODE.SNGL), MODE.MULT, LARGE, SMALL, RANK (=RANK.EQ), RANK.AVG, STDEV (=STDEV.S), STDEVP
  (=STDEV.P), VAR (=VAR.S), VARP (=VAR.P), STDEVA, STDEVPA, VARA, VARPA, PERCENTILE, QUARTILE
  and their .INC/.EXC variants, PERCENTRANK(.INC/.EXC), FREQUENCY, PROB, TRIMMEAN, AVEDEV,
  DEVSQ, GEOMEAN, HARMEAN, KURT, SKEW, SKEW.P, FISHER, FISHERINV, SUBTOTAL, AGGREGATE.
- **Correlation and regression**: CORREL, PEARSON, RSQ, COVAR (=COVARIANCE.P), COVARIANCE.S,
  SLOPE, INTERCEPT, STEYX, FORECAST (=FORECAST.LINEAR), LINEST, LOGEST, TREND, GROWTH.
- **Hypothesis tests**: TTEST (=T.TEST), FTEST (=F.TEST), CHITEST (=CHISQ.TEST), ZTEST (=Z.TEST).
- **Engineering**: BIN2DEC/HEX/OCT, DEC2…, HEX2…, OCT2… (ten digits, two's complement),
  DELTA, GESTEP, CONVERT (weight, length, time, pressure, force, energy, power, magnetism,
  temperature, volume, speed, bits and bytes — with decimal and binary prefixes), complex COMPLEX
  and IM-family (17 functions; `i`/`j` suffixes are not mixed), Bessel BESSELJ/Y/I/K (series and
  Hankel expansion, generated rather than tabulated).
- **Finance**: PMT, PV, FV, NPER, IPMT, PPMT, RATE, NPV, IRR, XNPV, XIRR, MIRR, SLN, SYD, DDB,
  DB, VDB, EFFECT, NOMINAL, CUMIPMT, CUMPRINC, ISPMT, FVSCHEDULE, DOLLARDE, DOLLARFR,
  AMORLINC, AMORDEGRC.
- **Securities**: coupon calendar COUPPCD/COUPNCD/COUPNUM/COUPDAYS/COUPDAYBS/COUPDAYSNC (with end-of-month
  sticking), PRICE, YIELD (bisection on price), DURATION, MDURATION, DISC, PRICEDISC, YIELDDISC,
  INTRATE, RECEIVED, PRICEMAT, YIELDMAT, TBILLPRICE, TBILLYIELD, TBILLEQ, ACCRINT (quasi-coupon
  periods), ACCRINTM, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD.
- **Database**: DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DPRODUCT, DGET, DSTDEV, DSTDEVP, DVAR,
  DVARP.
- **Arrays**: SORT, SORTBY, FILTER, UNIQUE, SEQUENCE, RANDARRAY, TRANSPOSE, SINGLE (`@`), VSTACK,
  HSTACK, TOCOL, TOROW, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, WRAPROWS, WRAPCOLS, EXPAND.
- **Distributions**: NORM.DIST/INV, NORM.S.DIST/INV, LOGNORM.DIST/INV, T.DIST/.2T/.RT, T.INV/.2T,
  CHISQ.DIST/.RT, CHISQ.INV/.RT, F.DIST/.RT, F.INV/.RT, BETA.DIST/INV, GAMMA.DIST/INV, EXPON.DIST,
  WEIBULL.DIST, POISSON.DIST, BINOM.DIST, BINOM.INV, NEGBINOM.DIST, HYPGEOM.DIST, plus GAMMA,
  GAMMALN(.PRECISE), ERF(.PRECISE), ERFC(.PRECISE), PHI, GAUSS, STANDARDIZE, CONFIDENCE(.NORM/.T).
  Pre-2010 names — `NORMDIST`, `TDIST`, `CHIINV`, `FINV`, `CRITBINOM`, and the other two dozen —
  are accepted alongside new ones: thirty-year-old files are full of them.
- **Honest stubs**: functions whose answer lives outside the file — the CUBE family and RTD without
  a data connection, GETPIVOTDATA without a pivot model — are registered and answer `#N/A`
  (GETPIVOTDATA — `#REF!`), as Excel does with a disconnected source; they do not yield `#NAME?`
  because the function is known.

An unknown function yields `#NAME?`, not an exception, so a file with exotic functions still opens
and recalculates everywhere else.

## Finance and database

Seven annuity functions are one identity rewritten for different unknowns:

```
pv·(1+r)ⁿ + pmt·(1+r·type)·((1+r)ⁿ − 1)/r + fv = 0
```

`PMT`, `PV`, `FV`, and `NPER` solve it algebraically; `IPMT` and `PPMT` split one payment into
interest and principal; `RATE` and `IRR` search for the rate numerically: Newton's method from a
hint (default 0.1) first, then bisection if it diverges. A rate at which cash flows converge may
not exist at all; then the answer is `#NUM!`, not a silent zero. `NPV` discounts from the first
period, `IRR` from zero, so initial outlays are written before the function in the first case and
included in the range in the second.

Database functions select records with a criteria table: criteria on one row must all match;
criteria on different rows are alternatives. There is no other way to express AND/OR to them. A
criteria column with nothing written is not checked at all — so one criteria range can be reused
for several queries. A field is named by column header or by 1-based column number; for `DCOUNT`
and `DCOUNTA` it may be omitted (or zero passed), and then records themselves are counted. `DGET`
requires exactly one matching record: none — `#VALUE!`; several — `#NUM!`.

Computed criteria (a formula under an empty or foreign header) are not supported: such a header is
not found among table columns, and the query answers `#VALUE!`.

`XNPV` and `XIRR` work on dates, not equal periods: a year is exactly 365 days, and everything is
measured from the first date in the list, so an earlier date is `#NUM!`. `MIRR` is the only member
of this family that needs no search: ordinary internal rate silently assumes withdrawn amounts are
reinvested at the same rate, while modified IRR asks separately what borrowed money costs and what
proceeds earn, and is solved by formula.

Of four depreciation methods, two declining ones (`DDB` and `DB`) are defined by period — each
charge is taken from what the previous one left — but are computed here by closed form, not by
walking years: the balance falls geometrically until it hits salvage and stops. Otherwise
`DDB(cost, salvage, 10*365, 1)` with life in days would cost three and a half thousand steps for
one number. The `DB` rate is additionally rounded to three digits, so `DB` and `DDB` diverge even
where their rates would otherwise match.

## Distributions

Each distribution is a thin read of one of four mathematical functions: incomplete gamma (normal
via error function, chi-square, Poisson), incomplete beta (Student, F, binomial, negative binomial),
log-gamma, and rational approximation for the normal quantile. Accuracy of the whole catalog rests
on them, so incomplete functions use a convergent expansion nearby: series where it converges fast,
continued fraction beyond.

The upper tail is computed with its own expansion, not by subtracting the lower tail from one. The
difference shows immediately: `CHISQ.DIST.RT(60,10)` is 3.6·10⁻⁹, and subtracting two numbers
near one would leave no correct digits. Inverse functions without a closed form are found by
bisection and checked in tests against their own forward function.

Excel renamed this catalog in 2010 — `NORMDIST` became `NORM.DIST`, and so on — but old names
still work, and both spellings register to one implementation. One trap is worth remembering:
`CHIDIST` and `CHISQ.DIST` are not the same function. The old name gives the right tail; the new
one gives the left.

## What makes recalculation fast

Three things worth knowing if the engine is used on large workbooks:

- **Read-range cache.** Formulas reading the same rectangle share one materialized copy: a column
  of 50,000 `VLOOKUP` against one table materializes it once, not per row. The cache lives between
  `Calculate` calls and is invalidated precisely when anything is written inside the rectangle —
  whether by the engine or by editing through the model. So a lookup table can live on the same
  sheet where results are written.
- **Range subscription index.** Editing a cell must find formulas that read it through a range.
  Identical ranges are merged into one subscription; the rest are laid out by column with an
  interval tree on rows, so lookup does not depend on the number of formulas in the workbook.
- **Exact-lookup index.** For a cached range of several hundred rows, `VLOOKUP`, `MATCH`, and
  `XLOOKUP` with exact match use a hash table instead of scanning. Approximate lookup, `*`/`?`
  patterns, and ranges with empty cells or errors stay on scan: their comparison does not reduce to
  equality that can be hashed.

## v1 limitations

- **Dynamic arrays spill.** A formula that returns an array occupies its rectangle: the anchor shows
  the first value, neighbors show the rest, `Cell.SpillRange` names the rectangle, `Cell.IsSpilled`
  distinguishes a spilled cell. An occupied cell, foreign spill, or merge in the way yields `#SPILL!`
  entirely — a truncated array would read as an answer. Reference `A1#` covers the whole rectangle
  and follows it on recalculation; writing to a spilled cell breaks the spill, and the anchor answers
  `#SPILL!` on the next recalculation — as in Excel. On save the anchor goes out as an array formula
  with `ref` and a `cm` index in metadata (`XLDAPR`), so Excel opens it as dynamic, and reload does
  not turn it into a CSE relic. Limits: rectangles are known only after evaluation, so formulas
  reading a spill may need an extra pass — the engine runs up to three per `Calculate` call; a bare
  range reference (`=A1:A3`) still collapses by implicit intersection, not spill; in a workbook that
  arrived with its own metadata part, a new spill is written as a legacy array with a
  `DynamicArrayAsLegacy` warning — the writer does not renumber someone else's part.
- No R1C1 input.
- `TEXT` renders culture-independently, like all other parsing: separators and month names do not
  depend on the machine running the code, whereas Excel takes them from the file locale. Fraction
  formats (`# ?/?`, `?/???`, `# ??/16`) are supported: the space allotted to the denominator is the
  whole assignment, and the nearest fraction under that ceiling is found by continued fractions.
- The "row hidden by filter" signal is inferred from AutoFilter criteria presence, not stored per
  row: a row hidden manually inside a filtered range counts as filtered. Criteria in `extLst` (x14)
  are not parsed at all.
- Arguments are evaluated eagerly: `IF` chooses a ready branch. The result matches Excel because
  functions are pure, but the unused branch is still computed.
- Dynamic dependencies from `INDIRECT`/`OFFSET` do not enter the graph: such formulas are always
  volatile and read current values. Order may lag by one pass in rare cases.
- No iterative cycle recalculation or user-defined functions.
- A reference to another workbook is evaluated from the `externalLink` cache — see above. From `.xls`
  such a reference is read with the full workbook path (`'[C:\…\file.xls]Sheet1'!A1`), but the old
  format does not give this library cached values, and the answer there remains `#NAME!` — see
  [xls.md](xls.md).

Known Excel divergences the oracle sees that are not yet closed:

- **`SUBTOTAL`** diverges where a filter that hid rows is no longer described in the file.

Two divergences from this list are closed. An error passed to `SUBTOTAL` and `AGGREGATE` as a whole
argument propagates outward — `SUBTOTAL(2,#REF!)` answers `#REF!`, as any call would through which
the evaluator would thread the error; `AGGREGATE` with the "ignore errors" option still ignores.
A `TEXT` format typed for another locale — `# ##0,00` where comma is the decimal separator and
space separates groups — parses correctly if you name the file locale in
`CalculateOptions.FormatCulture`: format and output separators come from it, while parsing stays
invariant by default. Localized date letters (`DD.MM`) are not translated.

## How correctness is verified

- **Oracle on real files** — the main signal. About 1100 xlsx files in sibling repositories store
  Excel-calculated values next to formulas. The test loads them, recalculates, and compares: currently
  399 files read, ~1550 formulas compared, 96.8% match, and results with unknown functions dropped
  from 99 to 9 — almost all contain add-in functions (`_xll.…`). Match rate grows more slowly than
  function count, and that is expected: each closed function moves a formula from "unsupported" to
  "compared", so the denominator grows with the numerator. The report prints functions by mismatch
  count with an example each, and serves as a priority list.
  The same oracle on the .xls corpus is two orders of magnitude wider — 54 thousand formulas, 98.2%
  match — and is described in [xls.md](xls.md) together with formula-text comparison against NPOI.
- **Parser corpus**: the same files, ~1560 formulas — parse without crashes, `parse → render → parse`
  is stable.
- **Fuzz**: 20,000 random strings; the only allowed exception is `FormulaParseException`.
- **Golden tests** for every semantics bullet above and every function.
