using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Conq.Noble.Spreadox.Tests;

public class PdfReportTests
{
    private static readonly byte[] TinyPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
        0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x03, 0x00, 0x01, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
    };

    [Fact]
    public void ToByteArray_WritesPdfHeaderAndCoverText()
    {
        using PdfReport report = PdfReport.Create();
        report.Cover.Title = "Analyst native PDF";
        report.Cover.Company = "Conq.Noble";
        report.Cover.Date = new DateTime(2026, 3, 15, 9, 30, 0);
        report.Cover.Provider = "BoQP";
        report.Cover.AddFact("Tag", "Station-A mean 12.4");
        report.Footer.Text = "provider-logo";
        report.Footer.PageNumberFormat = "Page {0} / {1}";
        report.Footer.Logo = PdfImage.FromBytes(TinyPng);
        report.Cover.Logo = PdfImage.FromBytes(TinyPng);

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] values = { 10, 12, 9, 15, 14, 13 };
        PdfChartPage page = report.AddChartPage("Station 12");
        page.Chart.Type = ExcelChartType.LineMarker;
        page.Chart.AddSeries("Actual", months, values).SetColor(ExcelColor.FromRgb(0, 82, 147));
        page.Chart.AddConstantLine("UCL", 16, ExcelColor.Red);
        page.Chart.AddConstantLine("CL", 12, ExcelColor.Green);
        page.Summary.Add("Mean", "12.2");
        page.Summary.Add("Max", "15");

        byte[] pdf = report.ToByteArray();
        AssertPdfHeader(pdf);
        using PdfDocument document = PdfDocument.Open(pdf);
        Assert.Equal(2, document.NumberOfPages);

        string cover = PageText(document.GetPage(1));
        Assert.Contains("Analyst native PDF", cover);
        Assert.Contains("Conq.Noble", cover);
        Assert.Contains("Station-A", cover);
        Assert.Contains("Charts:", cover);
        Assert.Contains("Page 1 / 2", cover);

        string chart = PageText(document.GetPage(2));
        Assert.Contains("Station 12", chart);
        Assert.Contains("Mean", chart);
        Assert.Contains("Page 2 / 2", chart);
    }

    [Fact]
    public void EmptySeries_WritesNoChartData()
    {
        using PdfReport report = PdfReport.Create();
        report.Cover.Title = "Empty";
        PdfChartPage page = report.AddChartPage("Blank chart");
        page.Chart.Type = ExcelChartType.ColumnClustered;

        byte[] pdf = report.ToByteArray();
        using PdfDocument document = PdfDocument.Open(pdf);
        string text = PageText(document.GetPage(2));
        Assert.Contains("No chart data", text);
    }

    [Fact]
    public void ComboPieBarAndStackedFallback_ProducePages()
    {
        using PdfReport report = PdfReport.Create();
        report.Cover.Title = "Mix";

        string[] cats = { "A", "B", "C", "D" };
        double[] bars = { 20, 35, 18, 27 };
        double[] line = { 40, 55, 48, 70 };
        double[] pie = { 30, 20, 50 };

        PdfChartPage combo = report.AddChartPage("Pareto");
        combo.Chart.Type = ExcelChartType.Combo;
        combo.Chart.AddSeries("Count", cats, bars).AsChartType(ExcelChartType.ColumnClustered);
        combo.Chart.AddSeries("Cum %", cats, line).AsChartType(ExcelChartType.Line).OnSecondaryAxis();
        combo.Chart.SecondaryValueAxis.Minimum = 0;
        combo.Chart.SecondaryValueAxis.Maximum = 100;
        combo.Chart.SecondaryValueAxis.Title = "%";

        PdfChartPage piePage = report.AddChartPage("Share");
        piePage.Chart.Type = ExcelChartType.Pie;
        piePage.Chart.AddSeries("Share", new[] { "North", "South", "East" }, pie);

        PdfChartPage stacked = report.AddChartPage("Stacked as bar");
        stacked.Chart.Type = ExcelChartType.ColumnStacked100;
        stacked.Chart.AddSeries("S1", cats, bars);

        PdfChartPage box = report.AddChartPage("Box as line");
        box.Chart.Type = ExcelChartType.BoxAndWhisker;
        box.Chart.AddSeries("Raw", cats, bars);

        PdfChartPage area = report.AddChartPage("Area");
        area.Chart.Type = ExcelChartType.Area;
        area.Chart.AddSeries("Fill", cats, bars);

        byte[] pdf = report.ToByteArray();
        AssertPdfHeader(pdf);
        using PdfDocument document = PdfDocument.Open(pdf);
        Assert.Equal(6, document.NumberOfPages);
        Assert.Contains("Pareto", PageText(document.GetPage(2)));
        Assert.Contains("Share", PageText(document.GetPage(3)));
        Assert.Contains("North", PageText(document.GetPage(3)));
    }

    [Fact]
    public void Summary_CapsAt28Rows()
    {
        using PdfReport report = PdfReport.Create();
        PdfChartPage page = report.AddChartPage("Stats");
        page.Chart.AddSeries("V", new[] { "A" }, new[] { 1.0 });
        for (int i = 0; i < 40; i++)
        {
            page.Summary.Add("M" + i, i.ToString());
        }

        Assert.Equal(PdfSummaryList.MaxRows, page.Summary.Count);
    }

    [Fact]
    public void MaxPoints_TruncatesSeries()
    {
        using PdfReport report = PdfReport.Create();
        var values = new double[8000];
        for (int i = 0; i < values.Length; i++)
        {
            values[i] = i;
        }

        PdfChartPage page = report.AddChartPage("Long");
        page.Chart.MaxPoints = 5000;
        PdfChartSeries series = page.Chart.AddSeries("N", values);
        Assert.Equal(5000, series.Values.Count);

        byte[] pdf = report.ToByteArray();
        AssertPdfHeader(pdf);
    }

    [Fact]
    public void Dispose_PreventsSave()
    {
        PdfReport report = PdfReport.Create();
        report.Dispose();
        Assert.Throws<ObjectDisposedException>(() => report.AddChartPage("X"));
        Assert.Throws<ObjectDisposedException>(() => report.ToByteArray());
    }

    [Fact]
    public void PdfTypes_StayInSpreadoxNamespace()
    {
        Assert.Equal("Conq.Noble.Spreadox", typeof(PdfReport).Namespace);
        Assert.Equal("Conq.Noble.Spreadox", typeof(ExcelWorkbook).Namespace);
        Assert.Equal("Conq.Noble.Spreadox", typeof(PdfHeader).Namespace);
        Assert.Equal("Conq.Noble.Spreadox", typeof(PdfFontRegistry).Namespace);
    }

    [Fact]
    public void ToByteArray_WithCustomHeader_RendersHeaderContent()
    {
        using PdfReport report = PdfReport.Create();
        report.Header.Title = "Custom Header Title";
        report.Header.Text = "Custom Subtitle Text";
        report.Header.Bold = true;
        report.Header.FontSize = 14;
        report.Header.TitleColor = ExcelColor.FromRgb(0, 82, 147);
        report.Header.Logo = PdfImage.FromBytes(TinyPng);

        report.Footer.Text = "Custom Footer Notice";
        report.Footer.PageNumberFormat = "P. {0} / {1}";
        report.Footer.Logo = PdfImage.FromBytes(TinyPng);

        PdfChartPage page = report.AddChartPage("Chart 1");
        page.Chart.AddSeries("S1", new[] { 10.0, 20.0, 30.0 });

        byte[] pdf = report.ToByteArray();
        AssertPdfHeader(pdf);

        using PdfDocument doc = PdfDocument.Open(pdf);
        Assert.Equal(2, doc.NumberOfPages);

        string page1Text = PageText(doc.GetPage(1));
        Assert.Contains("Custom Header Title", page1Text);
        Assert.Contains("Custom Subtitle Text", page1Text);
        Assert.Contains("Custom Footer Notice", page1Text);

        string page2Text = PageText(doc.GetPage(2));
        Assert.Contains("Custom Header Title", page2Text);
        Assert.Contains("Custom Subtitle Text", page2Text);
        Assert.Contains("Custom Footer Notice", page2Text);
    }

    private static void AssertPdfHeader(byte[] pdf)
    {
        Assert.True(pdf.Length > 8);
        string header = Encoding.ASCII.GetString(pdf, 0, 5);
        Assert.Equal("%PDF-", header);
    }

    private static string PageText(Page page)
    {
        return string.Join(" ", page.GetWords().Select(w => w.Text));
    }
}
