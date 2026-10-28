# Spreadox — Getting started

High-level Excel (`.xlsx`) API on `DocumentFormat.OpenXml` (`netstandard2.0`). OpenXml types stay **internal**.

**Package / namespace:** `Conq.Noble.Spreadox`

## Install

```bash
dotnet add package Conq.Noble.Spreadox
```

Local pack (from this repo):

```bash
dotnet pack src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj -c Release -o artifacts
# artifacts/Conq.Noble.Spreadox.0.7.0.nupkg
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
```
