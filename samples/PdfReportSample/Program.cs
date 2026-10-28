using Conq.Noble.Spreadox;

string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SpreadoxSamples");
Directory.CreateDirectory(outputDir);
string path = Path.Combine(outputDir, "analyst-native-pdf-sample.pdf");

string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
double[] actual = { 12.4, 11.8, 13.1, 12.0, 14.2, 13.6, 12.9, 15.0, 14.4, 13.2, 12.7, 13.8 };
double[] target = { 12, 12, 12, 12, 13, 13, 13, 14, 14, 13, 13, 13 };
double[] share = { 28, 22, 18, 15, 10, 7 };
string[] regions = { "North", "South", "East", "West", "Central", "Other" };
double[] counts = { 40, 28, 18, 10 };
double[] cumulative = { 40, 68, 86, 100 };
string[] pareto = { "A", "B", "C", "D" };

using var report = PdfReport.Create(PdfPageSize.A4Landscape);
report.Margins = PdfMargins.Millimeters(10, 10, 10, 14);
report.Cover.Company = "Conq.Noble";
report.Cover.Title = "Analyst native PDF";
report.Cover.Date = DateTime.Now;
report.Cover.Provider = "Spreadox sample";
report.Cover.AddFact("Tag", "Station-A  mean 13.1  stdev 0.9");
report.Cover.AddFact("Points", "12 months");
report.Footer.Text = "Spreadox";
report.Footer.PageNumberFormat = "Page {0} / {1}";

PdfChartPage linePage = report.AddChartPage("Station measurements");
linePage.Chart.Type = ExcelChartType.LineMarker;
linePage.Chart.AddSeries("Actual", months, actual).SetColor(ExcelColor.FromRgb(0, 82, 147));
linePage.Chart.AddConstantLine("UCL", 15.5, ExcelColor.FromRgb(200, 40, 40));
linePage.Chart.AddConstantLine("CL", 13.0, ExcelColor.FromRgb(0, 128, 0));
linePage.Chart.AddConstantLine("LCL", 10.5, ExcelColor.FromRgb(200, 40, 40));
linePage.Chart.CategoryAxis.LabelRotation = -35;
linePage.Summary.Add("Mean", "13.1");
linePage.Summary.Add("Max", "15.0");
linePage.Summary.Add("Min", "11.8");
linePage.Summary.Add("UCL", "15.5");

PdfChartPage comboPage = report.AddChartPage("Pareto");
comboPage.Chart.Type = ExcelChartType.Combo;
comboPage.Chart.AddSeries("Count", pareto, counts).AsChartType(ExcelChartType.ColumnClustered);
comboPage.Chart.AddSeries("Cum %", pareto, cumulative).AsChartType(ExcelChartType.Line).OnSecondaryAxis();
comboPage.Chart.LegendPosition = ExcelLegendPosition.Bottom;
comboPage.Summary.Add("Top", "A 40%");

PdfChartPage piePage = report.AddChartPage("Region share");
piePage.Chart.Type = ExcelChartType.Pie;
piePage.Chart.AddSeries("Share", regions, share);
piePage.Summary.Add("Largest", "North");

PdfChartPage areaPage = report.AddChartPage("Area vs target");
areaPage.Chart.Type = ExcelChartType.Area;
areaPage.Chart.AddSeries("Actual", months, actual);
areaPage.Chart.AddSeries("Target", months, target).SetColor(ExcelColor.FromRgb(237, 125, 49));

PdfChartPage emptyPage = report.AddChartPage("Empty");
emptyPage.Chart.Type = ExcelChartType.ColumnClustered;

report.Save(path);
Console.WriteLine("Wrote " + path);
