# Spreadox — Getting started

High-level Excel (`.xlsx`) API on `DocumentFormat.OpenXml` (`netstandard2.0`), plus A4 PDF chart reports. OpenXml types stay **internal**. PDF charts use in-memory points, not worksheet ranges.

**Package / namespace:** `Conq.Noble.Spreadox`

## Install

```bash
dotnet add package Conq.Noble.Spreadox
```

Local pack (from this repo):

```bash
dotnet pack src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj -c Release -o artifacts
# artifacts/Conq.Noble.Spreadox.0.8.0.nupkg
```

Or project-reference `src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj`.

## Workbook + chart (DOM)

Use `ExcelWorkbook` when you need random cell access, styles, merges, formulas, charts, or round-trip `Open`/`Save`.

```csharp
using Conq.Noble.Spreadox;

using var wb = ExcelWorkbook.Create();
var sheet = wb.Worksheets[0];
sheet.Name = "Measurements";
sheet["A1"].SetValue("Station");
sheet["B1"].SetValue("Actual");
// ... fill rows ...
var chart = sheet.Charts.Add(ExcelChartType.Combo, "Station measurements");
chart.AddSeries("B2:B9", "A2:A9", name: "Actual").AsChartType(ExcelChartType.ColumnClustered);
chart.AddSeries("C2:C9", "A2:A9", name: "Target").AsChartType(ExcelChartType.Line).OnSecondaryAxis();
wb.Save(@"C:\temp\report.xlsx");
```

Full runnable sample: `samples/MeasurementChartSample`.

## PDF report

```csharp
using Conq.Noble.Spreadox;

using var report = PdfReport.Create();
report.Cover.Title = "Measurements";
var page = report.AddChartPage("Station");
page.Chart.Type = ExcelChartType.LineMarker;
page.Chart.AddSeries("Actual", new[] { "A", "B" }, new[] { 1.0, 2.0 });
page.Summary.Add("Mean", "1.5");
report.Save(@"C:\temp\report.pdf");
```

Runnable sample: `samples/PdfReportSample`.

## Streaming (large sheetData)

Use `ExcelStreamingWorkbook` for 100k+ row exports. Rows are not kept in memory; charts are not supported on this path.

```csharp
using Conq.Noble.Spreadox;

using var sw = ExcelStreamingWorkbook.Create(@"C:\temp\large.xlsx");
var sheet = sw.AddWorksheet("Rows");
for (int r = 1; r <= 100_000; r++)
    sheet.WriteRow(r, row => { row.Text(1, "R" + r); row.Number(2, r); });
```

## Build & test

```bash
dotnet test Conq.Noble.Spreadox.sln -c Release --filter Category!=Perf
dotnet test Conq.Noble.Spreadox.sln -c Release --filter Category=Perf
# Skip 500k: SPREADOX_SKIP_500K=1
```

## Sample

```bash
dotnet run --project samples/MeasurementChartSample -c Release
dotnet run --project samples/FullDemo.NetCore -c Release
dotnet run --project samples/PdfReportSample -c Release
```
