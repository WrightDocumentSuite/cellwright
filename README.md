# Cellwright

[![CI](https://github.com/WrightDocumentSuite/cellwright/actions/workflows/ci.yml/badge.svg)](https://github.com/WrightDocumentSuite/cellwright/actions/workflows/ci.yml)

Современная .NET 10 библиотека для xlsx (Apache-2.0). Слоёная архитектура в духе `System.Text.Json`:
примитивы → стриминг → DOM → типизированный маппинг через source generator, плюс отдельный
пакет с формульным движком.

## Возможности v1

- **Стриминг с константной памятью**: `XlsxWriter` (forward-only, UTF-8 напрямую,
  `params ReadOnlySpan<CellValue>`) и `XlsxReader` (`IAsyncEnumerable<RowData>`).
- **DOM**: `Workbook`/`Worksheet` с data-oriented хранением — колоночные слайсы
  страницами по 512 строк, ни одного объекта на ячейку. Значения, стили, формулы
  (round-trip текстом), merge, freeze panes, скрытые листы.
- **Полнота листа**: скрытые строки и столбцы, группировка, автофильтр, гиперссылки,
  примечания, таблицы Excel, проверка данных, условное форматирование, настройки печати,
  защита листа и книги, цвет ярлыка и масштаб.
- **Структурная правка**: вставка и удаление строк и столбцов со сдвигом всего, что к ним
  привязано, и перенацеливанием формул всей книги; перестановка и копирование листов.
- **Свойства документа**: `Title`, `Author`, `Company`, даты и произвольные поля — то, что
  Excel показывает во вкладке «Сведения», а книга, созданная с нуля, раньше не несла вовсе.
- **Правка чужих файлов без потерь**: диаграммы, сводные таблицы, картинки, темы, макросы
  `.xlsm` переносятся в сохранённый файл как есть; переименование листа тянет за собой
  формулы, имена и ссылки диаграмм.
- **Стили**: immutable records + бесплатное интернирование; fluent-строители на
  C# 14 extension members: `CellStyle.Default.Bold().Fill(0xFF2278D4).NumberFormat("#,##0.00")`.
  Покрыт весь §18.8: узорные и градиентные заливки, диагонали, поворот текста, цвета темы,
  защита ячеек, стили строк и столбцов целиком; именованные стили ячеек и `dxf`-форматы
  загруженного файла сохраняются и читаются.
- **`CellValue`** — 24-байтный tagged union без боксинга; даты с тиковой точностью,
  честные конверсии серийных дат (1900-баг, система 1904).
- **Типизированный маппинг** (`Cellwright.Mapping`): `[ExcelRow]` + incremental source
  generator → `IRowSchema<T>` (static abstract members), zero reflection, Native AOT.
- **Формульный движок** (`Cellwright.Formulas`): 447 функций — 352 из 356 по §18.17.7
  ISO/IEC 29500 (без четырёх японо-тайских), включая ценные бумаги, инженерные, комплексные,
  Бесселя, корреляцию и тесты гипотез, — инкрементальный пересчёт по графу зависимостей,
  детекция циклов, разворачивание shared-формул, defined names, структурные ссылки
  `Таблица[Колонка]`, 3-D диапазоны `Q1:Q4!A1`, ссылки в закрытые книги из кэша
  `externalLink`, `LET` и `TEXT`.
- **Legacy `.xls`** (`Cellwright.Xls`): чтение BIFF8, BIFF5 и BIFF2–4 в тот же `Workbook`,
  запись BIFF8 (включая примечания через слой Office Art, проверки данных и настройки печати),
  расшифровка XOR/RC4/RC4 CryptoAPI, декодирование и кодирование формул Ptg, чтение условного
  форматирования и картинок из слоя Escher. Конвертация xls → xlsx получается сама собой —
  вместе с картинками, — а `XlsReader` отдаёт строки потоком, если модель не нужна.
- **Двоичный `.xlsb`** (`Cellwright.Xls`): чтение записей BIFF12 в тот же `Workbook` — значения,
  стили, строки, формулы со структурными ссылками, примечания, проверка данных, условное
  форматирование, таблицы и картинки, включая зашифрованные книги. Записи `.xlsb` нет: сохранять
  в xlsx или `.xls`.
- **Картинки и тема**: изображения читаются из всех трёх форматов и пишутся в xlsx, байты не
  пережимаются; `xl/theme/theme1.xml` разбирается, так что цвет темы разрешается в конкретный
  ARGB вместе с тинтом — см. [pictures.md](docs/pictures.md).
- **Пароль на открытие**: зашифрованный `.xlsx` и `.xlsb` — тот же контейнер MS-CFB, поэтому их
  читает `Cellwright.Xls`: обе схемы MS-OFFCRYPTO, шифры AES/DES/3DES/RC2, все пять хэшей,
  плюс проверка HMAC над пакетом. Пишется тоже — `SaveEncryptedAsync` кладёт книгу под agile
  AES-256/SHA-512, как это делает Excel; записанное открывают NPOI и ExcelDataReader.
- **Макросы**: исходники VBA извлекаются из `.xlsm` и `.xls` — распаковка MS-OVBA, виды модулей,
  подключённые библиотеки, состояние защиты. Пароль на проект исходник не прячет.
- **Экспорт в Markdown**: `book.ToMarkdown()` отдаёт документ GitHub Flavored Markdown — лист
  разделом с таблицей, значения отрендерены как в Excel, объединения, гиперссылки, картинки
  галереей; `ExportMarkdownAsync` кладёт `workbook.md` и `media/` в каталог. Чистое преобразование
  модель → текст, без бинарных зависимостей — см. [markdown.md](docs/markdown.md).
- **Печать в PDF** (`Cellwright.Pdf`): `book.SaveAsPdf("отчёт.pdf")` печатает книгу так, как её
  печатает Excel — сетка со стилями и числовыми форматами, объединения, перенос и вытекание текста,
  повороты, картинки по якорям, двумерная разбивка на страницы с масштабом и `FitToWidth`,
  повторяющиеся заголовки, колонтитулы с `&P`/`&N`/`&A`. Опционально с деревом структуры (лист —
  таблица, ячейки — `TD`/`TH`). Ничего не запускается снаружи, шрифты не подменяются молча —
  каждая подмена названа в диагностике; см. [pdf-export.md](docs/pdf-export.md).
- **Диагностика вместо тихих отказов**: recovery-ориентированная загрузка,
  `WorkbookWarning` с кодами, `LoadDiagnostics`; файл не того формата (`.ods`, `.docx`,
  зашифрованный пакет) называется по имени, а не роняет парсер.
- **Чужой файл как враждебный ввод**: `LoadOptions.Limits` ограничивает, во что пакету
  позволено развернуться, — zip-бомба получает сообщение, а не `OutOfMemoryException`.
- **Async-first**: I/O пакета на асинхронных `ZipArchive` API .NET 10.

## Быстрый пример

```csharp
using Cellwright;

// DOM
var book = Workbook.Create();
Worksheet sheet = book.Sheets.Add("Продажи");
sheet["A1"] = "Товар";
sheet["B1"] = 1234.5;
sheet.Range("A1:B1").Style(s => s.Bold().Fill(0xFF2278D4));
sheet.Cell("B3").SetFormula("SUM(B1:B2)");
sheet.Column("A").SetWidth(24);
sheet.AutoFilter("A1:B3");
await book.SaveAsync("sales.xlsx");

// Правка реального отчёта: диаграммы, сводные и макросы остаются на месте
Workbook report = await Workbook.LoadAsync("отчёт.xlsm");
report.Sheets[0]["B7"] = 42;
await report.SaveAsync("отчёт-обновлён.xlsm");

// Стриминг миллионов строк
await using XlsxWriter writer = await XlsxWriter.CreateAsync("big.xlsx");
XlsxSheetWriter data = await writer.BeginSheetAsync("Data");
for (int i = 1; i <= 1_000_000; i++)
{
    data.WriteRow(i, "name-" + (i % 1000), DateTime.UtcNow, i % 2 == 0);
    await data.FlushIfNeededAsync();
}
```

```csharp
// Типизированные строки (Cellwright.Mapping, source generator)
[ExcelRow]
public partial record Order(
    [ExcelColumn("Order ID")] int Id,
    string Customer,
    [ExcelColumn(Format = "#,##0.00")] decimal Total,
    DateTime? ShippedAt);

await foreach (Order order in reader.ReadRowsAsync<Order>("Orders")) { }
await sheet.WriteRowsAsync(orders);
```

```csharp
// Формулы (Cellwright.Formulas)
sheet.Cell("D2").SetFormula("B2*C2*VLOOKUP(A2,Rates!$A$2:$B$9,2,FALSE)");
sheet.Cell("D3").SetFormula("LET(итог,SUM(Продажи[Сумма]),TEXT(итог,\"#,##0.00\"))");
sheet.Cell("D4").SetFormula("PMT(0.08/12,60,30000)");            // -608,29
sheet.Cell("D5").SetFormula("DSUM(База!A1:E99,\"Прибыль\",Условия!A1:B2)");
sheet.Cell("D6").SetFormula("SUM(FILTER(B2:B99,A2:A99=\"Север\"))");
sheet.Cell("D7").SetFormula("NORM.DIST(42,40,1.5,TRUE)");        // 0,908789
CalculationSummary summary = book.Calculate();   // первый раз считает всё
sheet["B2"] = 500;
book.Calculate();                                 // дальше — только зависимое
CellValue total = book.Evaluate("SUM(Report!D2:D9)");
```

```csharp
// Правка структуры: всё привязанное к ячейкам едет за ними, формулы книги перенацеливаются
sheet.InsertRows(2, 3);
sheet.DeleteColumns(5);
sheet.Name = "Q1 2026";                       // формулы, имена и ссылки диаграмм следуют за листом
book.Sheets.Copy("Q1 2026", "Q2 2026");       // копия со всем содержимым, сразу за оригиналом
book.Sheets.Move("Итоги", 0);

book.Properties.Title = "Отчёт за квартал";   // вкладка «Сведения» в Excel
book.Properties.Custom["Отдел"] = CellValue.Text("Финансы");
```

```csharp
// Legacy .xls (Cellwright.Xls): читается в тот же Workbook, сохраняется куда угодно
Workbook legacy = await XlsWorkbook.LoadAsync("отчёт-2003.xls");
await legacy.SaveAsync("отчёт-2026.xlsx");            // конвертация из коробки
await legacy.SaveXlsAsync("отчёт-2003-обновлён.xls"); // или обратно в бинарный

Workbook any = await Spreadsheet.LoadAsync(uploaded); // xlsx, xls или xlsb — по содержимому

// Двоичный .xlsb читается тем же вызовом
Workbook binary = await Spreadsheet.LoadAsync("отчёт.xlsb");
await binary.SaveAsync("отчёт.xlsx");

// Пароль на открытие — что у бинарного файла, что у пакета
Workbook locked = await Spreadsheet.LoadAsync("закрытый.xlsx", new XlsLoadOptions { Password = "секрет" });

// Макросы: исходники и подключённые библиотеки из обоих форматов
VbaProject? macros = await VbaProject.LoadAsync("отчёт.xlsm");
foreach (VbaModule module in macros?.Modules ?? [])
    Console.WriteLine($"{module.Kind} {module.Name}: {module.Code.Length} символов");
foreach (VbaReference reference in macros?.References ?? [])
    Console.WriteLine($"{reference.Name}: {reference.Description}");

// Только значения, без модели
await using XlsReader rows = await XlsReader.OpenAsync("выгрузка.xls");
await foreach (RowData row in rows.ReadRowsAsync(0))
    Console.WriteLine(row.GetValueOrBlank(0));
```

## Производительность

100 000 строк × 20 колонок (2 млн ячеек), детали в [docs/benchmarks.md](docs/benchmarks.md):

| | Запись | Аллокации | Чтение | Аллокации |
| --- | ---: | ---: | ---: | ---: |
| **Cellwright (streaming)** | **3.5 s** | **32 MB** | **3.8 s** | **607 MB** |
| Cellwright (DOM) | 4.0 s | 78 MB | 5.6 s | 654 MB |
| ClosedXML 0.105 | 13.3 s | 2 095 MB | 25.5 s | 2 602 MB |
| EPPlus 8.6 | 16.3 s | 2 142 MB | 17.1 s | 1 782 MB |

Пересчёт 100 000 формул (VLOOKUP по справочнику, вложенные IF, итоги) — 112 мс; правка одной
ячейки в посчитанной книге — от 68 мкс.

Legacy `.xls`, 200 000 ячеек: чтение 51 мс против 142 мс у ExcelDataReader, аллокации
17.7 MB против 96.1 MB.

## Документация

- [Быстрый старт](docs/getting-started.md)
- [Архитектура](docs/architecture.md)
- [Стриминг](docs/streaming.md)
- [Лист: строки, столбцы и всё вокруг ячеек](docs/worksheet.md)
- [Стили](docs/styles.md)
- [Маппинг](docs/mapping.md)
- [Формулы](docs/formulas.md)
- [Legacy .xls](docs/xls.md)
- [Двоичный .xlsb](docs/xlsb.md)
- [Картинки и цвета темы](docs/pictures.md)
- [Экспорт в Markdown](docs/markdown.md)
- [Печать в PDF](docs/pdf-export.md)
- [Макросы (VBA)](docs/vba.md)
- [Соответствие спецификациям](docs/conformance.md)
- [Бенчмарки](docs/benchmarks.md)

## Сборка

Требуется .NET 10 SDK.

```bash
dotnet build Cellwright.slnx
dotnet test Cellwright.slnx
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks -- --filter * --inProcess
dotnet run --project samples/Cellwright.Samples
```

CI собирает и прогоняет тесты на Linux и Windows при каждом push и pull request
(`.github/workflows/ci.yml`). Корпус реальных книг лежит в соседних репозиториях, которых в
git нет, — зависящие от него тесты там пропускаются сами.

## Структура

```
src/Cellwright                    ядро: примитивы, OPC, стриминг, DOM, стили
src/Cellwright.Formulas           формульный движок: парсер, исполнитель, пересчёт
src/Cellwright.Pdf                печать в PDF поверх Inkwright: сетка, страницы, колонтитулы
src/Cellwright.Xls                двоичные форматы: CFB, BIFF8 (.xls), BIFF12 (.xlsb), Ptg-формулы,
                                  расшифровка, извлечение VBA
src/Cellwright.Mapping            атрибуты + runtime маппинга
src/Cellwright.Mapping.Generator  Roslyn incremental generator (пакуется в Mapping)
tests/                            xUnit: 1834 теста (прогоняются дважды, второй раз под культурой
                                  с запятой в дробях), включая OpenXml-валидатор, interop с ClosedXML,
                                  round-trip ~400 реальных xlsx и xlsm без потери частей, 844 реальных
                                  .xls и 66 .xlsb со сверкой значений, гиперссылок, свойств документа,
                                  картинок и цветов темы против ExcelDataReader и NPOI, сверку 64 книг,
                                  лежащих в корпусе сразу в двух форматах, пересчёт формул трёх форматов
                                  против значений Excel, расхождения Excel со стандартом по MS-OI29500,
                                  595 модулей VBA и 301 ссылка на библиотеки из 114 книг с макросами
benchmarks/                       BenchmarkDotNet vs EPPlus 8 / ClosedXML / ExcelDataReader
samples/                          рабочие примеры всех слоёв
```

Осознанные ограничения перечислены в [architecture.md](docs/architecture.md),
[worksheet.md](docs/worksheet.md), [formulas.md](docs/formulas.md), [xls.md](docs/xls.md),
[xlsb.md](docs/xlsb.md), [pictures.md](docs/pictures.md), [markdown.md](docs/markdown.md) и
[vba.md](docs/vba.md): нет AutoFit, нет API для диаграмм и сводных таблиц (они только переносятся
в сохранённый xlsx), нет spill у динамических массивов и нет `LAMBDA`; внешние книги считаются
из кэша связи, но сам чужой файл не открывается; макросы
только читаются, но не пишутся; шифрование пакета пишется, шифрование `.xls` — нет; `.xlsb`
читается, но не пишется; бинарный `.xls` не переносит непонятые части и не пишет условное
форматирование; картинки читаются везде, но пишутся только в xlsx; Markdown только пишется,
обратно не читается.
