using System.Data;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelPhase2Tests
{
    [Fact]
    public void Styles_NumberFormats_RoundTrip_AndValidate()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet.Name = "Styled";

            ExcelCell header = sheet["A1"];
            header.SetValue("Title");
            header.Style.Font.Bold = true;
            header.Style.Font.Name = "Calibri";
            header.Style.Font.Size = 14;
            header.Style.Font.Color = ExcelColor.White;
            header.Style.Fill.Pattern = ExcelFillPattern.Solid;
            header.Style.Fill.ForegroundColor = ExcelColor.FromRgb(0, 82, 147);
            header.Style.Alignment.Horizontal = ExcelHorizontalAlignment.Center;
            header.Style.Borders.SetAll(ExcelBorderStyle.Thin, ExcelColor.Black);

            sheet["B1"].SetValue(0.256);
            sheet["B1"].Style.NumberFormat = NumberFormats.PercentTwoDecimals;

            sheet["C1"].SetValue(1234.5);
            sheet["C1"].Style.NumberFormat = NumberFormats.ThousandsTwoDecimals;

            sheet["D1"].SetValue(new DateTime(2024, 1, 2));
            sheet["D1"].Style.NumberFormat = NumberFormats.Date;

            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelWorksheet s = loaded["Styled"];
        Assert.True(s["A1"].Style.Font.Bold);
        Assert.Equal("Calibri", s["A1"].Style.Font.Name);
        Assert.Equal(ExcelFillPattern.Solid, s["A1"].Style.Fill.Pattern);
        Assert.Equal(NumberFormats.PercentTwoDecimals, s["B1"].Style.NumberFormat);
        Assert.Equal(0.256, s["B1"].GetDouble()!.Value, 5);
        Assert.Equal(NumberFormats.Date, s["D1"].Style.NumberFormat);
    }

    [Fact]
    public void RowColumn_Merge_Freeze_AutoFilter_NamedRange()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet.Name = "Layout";
            sheet.SetRowHeight(1, 30);
            sheet.SetColumnWidth("A", 20);
            sheet.SetColumnWidth(2, 12);
            sheet["A1"].SetValue("Merged");
            sheet.Merge("A1:C1");
            sheet.FreezePanes(1, 1);
            sheet.AutoFilterRange = "A1:C3";
            sheet["A2"].SetValue(1);
            sheet["B2"].SetValue(2);
            sheet["C2"].SetValue(3);
            wb.AddNamedRange("LayoutData", "Layout!$A$2:$C$2");
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelWorksheet sheet2 = loaded["Layout"];
        Assert.True(sheet2.TryGetRowHeight(1, out double h) && Math.Abs(h - 30) < 0.01);
        Assert.True(sheet2.TryGetColumnWidth(1, out double w) && Math.Abs(w - 20) < 0.01);
        Assert.Contains(sheet2.MergedRanges, m => m.A1 == "A1:C1");
        Assert.Equal(1, sheet2.FreezeRows);
        Assert.Equal(1, sheet2.FreezeColumns);
        Assert.Equal("A1:C3", sheet2.AutoFilterRange);
        Assert.Contains(loaded.NamedRanges, n => n.Name == "LayoutData");
    }

    [Fact]
    public void Formula_Write_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(10);
            sheet["A2"].SetValue(20);
            sheet.SetFormula("A3", "A1+A2", cachedValue: 30);
            sheet["B1"].SetFormula("=SUM(A1:A2)");
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.Equal(CellValueKind.Formula, loaded.Worksheets[0]["A3"].Kind);
        Assert.Equal("A1+A2", loaded.Worksheets[0]["A3"].Formula);
        Assert.Equal(30, Convert.ToInt32(loaded.Worksheets[0]["A3"].Value));
        Assert.Equal("SUM(A1:A2)", loaded.Worksheets[0]["B1"].Formula);
    }

    [Fact]
    public void BulkInsert_DataTable_And_Objects()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet.Insert(new[]
            {
                new object?[] { "A", 1 },
                new object?[] { "B", 2 }
            }, startRow: 1, startColumn: 1);

            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Qty", typeof(int));
            table.Rows.Add("Widget", 5);
            table.Rows.Add("Gadget", 7);
            sheet.InsertDataTable(table, startRow: 4, startColumn: 1, includeHeaders: true);

            sheet.InsertObjects(
                new[]
                {
                    new SampleRow { Sku = "X1", Price = 1.5 },
                    new SampleRow { Sku = "X2", Price = 2.5 }
                },
                startRow: 8,
                startColumn: 1,
                includeHeaders: true);

            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelWorksheet s = loaded.Worksheets[0];
        Assert.Equal("A", s["A1"].GetString());
        Assert.Equal(2, s["B2"].GetInt());
        Assert.Equal("Name", s["A4"].GetString());
        Assert.Equal("Widget", s["A5"].GetString());
        Assert.Equal(7, s["B6"].GetInt());
        Assert.Equal("Sku", s["A8"].GetString());
        Assert.Equal("X2", s["A10"].GetString());
        Assert.Equal(2.5, s["B10"].GetDouble()!.Value, 5);
    }

    [Fact]
    public void ConditionalFormatting_WritesValidPackage()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(5);
            sheet["A2"].SetValue(15);
            sheet.AddConditionalFormat("A1:A2", ConditionalOperator.GreaterThan, "10", ExcelColor.Yellow);
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.NotEmpty(loaded.Worksheets[0].ConditionalFormats);
    }

    private sealed class SampleRow
    {
        public string Sku { get; set; } = "";
        public double Price { get; set; }
    }

    private static void AssertNoOpenXmlErrors(Stream stream)
    {
        using MemoryStream copy = new MemoryStream();
        long original = stream.CanSeek ? stream.Position : 0;
        stream.CopyTo(copy);
        if (stream.CanSeek)
        {
            stream.Position = original;
        }

        copy.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(copy, false);
        OpenXmlValidator validator = new OpenXmlValidator();
        List<ValidationErrorInfo> errors = validator.Validate(document).ToList();
        if (errors.Count == 0)
        {
            return;
        }

        var sb = new StringBuilder();
        sb.AppendLine($"OpenXmlValidator reported {errors.Count} error(s):");
        foreach (ValidationErrorInfo error in errors.Take(20))
        {
            sb.AppendLine($"- {error.Description} ({error.Path?.XPath})");
        }

        Assert.True(false, sb.ToString());
    }
}
