# Spreadox

[![NuGet](https://img.shields.io/nuget/v/Conq.Noble.Spreadox.svg)](https://www.nuget.org/packages/Conq.Noble.Spreadox)
[![CI](https://github.com/FatihBaycu/Conq.Noble.Spreadox/actions/workflows/ci.yml/badge.svg)](https://github.com/FatihBaycu/Conq.Noble.Spreadox/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

High-level C# library for Excel (`.xlsx`) workbooks **and** A4 PDF chart reports. Excel is built on Microsoft Open XML SDK (`DocumentFormat.OpenXml`). Public API stays free of OpenXml types. PDF charts are drawn as vectors from in-memory points — they are not OOXML charts.

| | |
|--|--|
| **Package** | `Conq.Noble.Spreadox` |
| **Namespace** | `Conq.Noble.Spreadox` |
| **Target** | `netstandard2.0` (.NET Framework 4.8+, .NET 6/8+) |
| **License** | MIT |

## Install

```bash
dotnet add package Conq.Noble.Spreadox
```

```csharp
using Conq.Noble.Spreadox;
```

## Quick usage

```csharp
using Conq.Noble.Spreadox;

using var wb = ExcelWorkbook.Create();
var sheet = wb.Worksheets[0];
sheet.Name = "Demo";
sheet["A1"].SetValue("hello");
sheet["A1"].Style.Font.Bold = true;
sheet["B1"].SetValue(0.25);
sheet["B1"].Style.NumberFormat = NumberFormats.Percent;
sheet.SetFormula("D1", "B1*100", cachedValue: 25);
sheet.FreezePanes(1);

var chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "Demo chart");
chart.AddSeries("B2:B10", "A2:A10", name: "Values");
chart.SetPosition(2, 4, 16, 12);

var combo = sheet.Charts.Add(ExcelChartType.Combo, "Combo");
combo.AddSeries("B2:B10", "A2:A10", name: "Bars").AsChartType(ExcelChartType.ColumnClustered);
combo.AddSeries("C2:C10", "A2:A10", name: "Line").AsChartType(ExcelChartType.Line).OnSecondaryAxis();
combo.SetPosition(2, 13, 16, 20);

wb.Save(@"C:\temp\demo.xlsx");

using (var sw = ExcelStreamingWorkbook.Create(@"C:\temp\large.xlsx"))
{
    var big = sw.AddWorksheet("Rows");
    for (int r = 1; r <= 100_000; r++)
    {
        big.WriteRow(r, row =>
        {
            row.Text(1, "R" + r);
            row.Number(2, r);
        });
    }
}
```

See `samples/MeasurementChartSample` and `samples/FullDemo.NetCore` for longer walkthroughs.

## PDF reports

`PdfReport` is a second surface in the same package. Use it for cover + one-chart-per-page A4 landscape PDFs (Analyst-style). Series are arrays, not cell ranges.

```csharp
using var report = PdfReport.Create(PdfPageSize.A4Landscape);
report.Cover.Title = "Analyst native PDF";
report.Cover.Company = "Conq.Noble";
report.Footer.PageNumberFormat = "Page {0} / {1}";

var page = report.AddChartPage("Station 12");
page.Chart.Type = ExcelChartType.LineMarker;
page.Chart.AddSeries("Actual", months, values).SetColor(ExcelColor.FromRgb(0, 82, 147));
page.Chart.AddConstantLine("UCL", 16, ExcelColor.Red);
page.Summary.Add("Mean", "12.2");

report.Save(@"C:\temp\report.pdf");
```

v1 chart kinds: Line+marker, Area, clustered Bar, Pie, Combo (Bar + Line on a 0–100 secondary axis), horizontal constant lines. `BoxAndWhisker` draws as line+marker; stacked-100 draws as clustered bar. Sample: `samples/PdfReportSample`.

## DOM vs streaming

| | `ExcelWorkbook` | `ExcelStreamingWorkbook` |
|--|-----------------|--------------------------|
| Cell model | In-memory grid | Rows flushed as written |
| Charts / pictures / CF / open | Yes | Write-only (no charts) |
| Large exports | OK to tens of k rows | Prefer for 100k+ rows |
| Strings | Shared string table | `inlineStr` (default) or SST |

## Solution layout

- `src/Conq.Noble.Spreadox` — library
- `tests/Conq.Noble.Spreadox.Tests` — xUnit + `OpenXmlValidator`
- `samples/` — net8 and net48 console demos, plus `PdfReportSample`
- `docs/getting-started.md` — short usage guide

## Build, test, pack

```bash
dotnet test Conq.Noble.Spreadox.sln -c Release --filter Category!=Perf
dotnet pack src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj -c Release -o artifacts
```

Optional performance suite:

```bash
dotnet test Conq.Noble.Spreadox.sln -c Release --filter Category=Perf
# Skip 500k: SPREADOX_SKIP_500K=1
```

Or: `./build-and-test.ps1`

A GitHub Release tagged `v0.8.0` (or later) publishes the package to nuget.org when the `NUGET_API_KEY` repository secret is set.

## License

MIT — see [LICENSE](LICENSE). Depends on `DocumentFormat.OpenXml` (MIT) and `PdfSharpCore` (MIT). Embedded Liberation Sans fonts are SIL Open Font License.
