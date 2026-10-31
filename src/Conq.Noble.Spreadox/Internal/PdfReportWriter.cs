using System;
using System.Globalization;
using System.IO;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace Conq.Noble.Spreadox.Internal
{
    internal static class PdfReportWriter
    {
        private const double HeaderHeight = 36;
        private const double FooterHeight = 22;

        public static void Write(PdfReport report, Stream stream)
        {
            PdfFontResolver.EnsureInstalled();

            using (var document = new PdfDocument())
            {
                document.Info.Title = report.Cover.Title ?? "Spreadox report";
                document.Info.Author = report.Cover.Company ?? "Spreadox";
                document.Info.Creator = "Conq.Noble.Spreadox";

                int chartCount = report.Cover.ChartCount ?? report.ChartPages.Count;
                int totalPages = 1 + report.ChartPages.Count;
                if (totalPages < 1)
                {
                    totalPages = 1;
                }

                DrawCover(document, report, chartCount, 1, totalPages);
                for (int i = 0; i < report.ChartPages.Count; i++)
                {
                    DrawChartPage(document, report, report.ChartPages[i], i + 2, totalPages);
                }

                document.Save(stream, closeStream: false);
            }
        }

        private static PdfPage AddPage(PdfDocument document)
        {
            PdfPage page = document.AddPage();
            page.Size = PdfSharpCore.PageSize.A4;
            page.Orientation = PdfSharpCore.PageOrientation.Landscape;
            return page;
        }

        private static bool HasHeaderContent(PdfHeader header)
        {
            return header != null && (!string.IsNullOrWhiteSpace(header.Title)
                || !string.IsNullOrWhiteSpace(header.Text)
                || header.Logo != null);
        }

        private static XRect ContentRect(PdfPage page, PdfMargins margins, bool hasHeader)
        {
            double left = Mm(margins.Left);
            double top = Mm(margins.Top) + (hasHeader ? HeaderHeight : 0);
            double right = Mm(margins.Right);
            double bottom = Mm(margins.Bottom) + FooterHeight;
            return new XRect(left, top, page.Width - left - right, page.Height - top - bottom);
        }

        private static void DrawCover(PdfDocument document, PdfReport report, int chartCount, int pageNo, int totalPages)
        {
            PdfPage page = AddPage(document);
            using (XGraphics gfx = XGraphics.FromPdfPage(page))
            {
                bool hasHeader = HasHeaderContent(report.Header);
                if (hasHeader)
                {
                    DrawHeader(gfx, page, report);
                }

                XRect content = ContentRect(page, report.Margins, hasHeader);
                var titleFont = new XFont(PdfFontResolver.FamilyName, 22, XFontStyle.Bold);
                var companyFont = new XFont(PdfFontResolver.FamilyName, 12, XFontStyle.Regular);
                var bodyFont = new XFont(PdfFontResolver.FamilyName, 10, XFontStyle.Regular);
                var factFont = new XFont(PdfFontResolver.FamilyName, 9, XFontStyle.Regular);
                var accent = new XSolidBrush(XColor.FromArgb(0, 82, 147));

                double y = content.Y;
                if (report.Cover.Logo != null)
                {
                    y += DrawImageAspectFit(gfx, report.Cover.Logo, new XRect(content.X, y, 120, 48));
                    y += 10;
                }

                if (!string.IsNullOrWhiteSpace(report.Cover.Company))
                {
                    gfx.DrawString(report.Cover.Company, companyFont, accent, new XRect(content.X, y, content.Width, 18), XStringFormats.TopLeft);
                    y += 20;
                }

                string title = string.IsNullOrWhiteSpace(report.Cover.Title) ? "Report" : report.Cover.Title!;
                gfx.DrawString(title, titleFont, XBrushes.Black, new XRect(content.X, y, content.Width, 32), XStringFormats.TopLeft);
                y += 36;

                gfx.DrawLine(new XPen(XColor.FromArgb(0, 82, 147), 1.5), content.X, y, content.X + 220, y);
                y += 14;

                if (report.Cover.Date.HasValue)
                {
                    gfx.DrawString(report.Cover.Date.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), bodyFont, XBrushes.DimGray,
                        new XRect(content.X, y, content.Width, 16), XStringFormats.TopLeft);
                    y += 18;
                }

                gfx.DrawString("Charts: " + chartCount.ToString(CultureInfo.InvariantCulture), bodyFont, XBrushes.Black,
                    new XRect(content.X, y, content.Width, 16), XStringFormats.TopLeft);
                y += 18;

                if (!string.IsNullOrWhiteSpace(report.Cover.Provider))
                {
                    gfx.DrawString(report.Cover.Provider, bodyFont, XBrushes.DimGray,
                        new XRect(content.X, y, content.Width, 16), XStringFormats.TopLeft);
                    y += 18;
                }

                y += 8;
                for (int i = 0; i < report.Cover.Facts.Count; i++)
                {
                    PdfCoverFact fact = report.Cover.Facts[i];
                    gfx.DrawString(fact.Label + ": " + fact.Value, factFont, XBrushes.Black,
                        new XRect(content.X, y, content.Width, 14), XStringFormats.TopLeft);
                    y += 14;
                }

                DrawFooter(gfx, page, report, pageNo, totalPages);
            }
        }

        private static void DrawChartPage(PdfDocument document, PdfReport report, PdfChartPage chartPage, int pageNo, int totalPages)
        {
            PdfPage page = AddPage(document);
            using (XGraphics gfx = XGraphics.FromPdfPage(page))
            {
                bool hasHeader = HasHeaderContent(report.Header);
                if (hasHeader)
                {
                    DrawHeader(gfx, page, report);
                }

                XRect content = ContentRect(page, report.Margins, hasHeader);
                var titleFont = new XFont(PdfFontResolver.FamilyName, 14, XFontStyle.Bold);
                gfx.DrawString(chartPage.Title, titleFont, XBrushes.Black,
                    new XRect(content.X, content.Y, content.Width, 22), XStringFormats.TopLeft);

                double summaryWidth = chartPage.Summary.Count > 0 ? 150 : 0;
                var chartBounds = new XRect(content.X, content.Y + 26, content.Width - summaryWidth - (summaryWidth > 0 ? 10 : 0), content.Height - 26);
                PdfChartPainter.Draw(gfx, chartBounds, chartPage.Chart);

                if (chartPage.Summary.Count > 0)
                {
                    DrawSummary(gfx, new XRect(content.Right - summaryWidth, content.Y + 26, summaryWidth, content.Height - 26), chartPage);
                }

                DrawFooter(gfx, page, report, pageNo, totalPages);
            }
        }

        private static void DrawSummary(XGraphics gfx, XRect bounds, PdfChartPage chartPage)
        {
            var headerFont = new XFont(PdfFontResolver.FamilyName, 9, XFontStyle.Bold);
            var bodyFont = new XFont(PdfFontResolver.FamilyName, 8, XFontStyle.Regular);
            gfx.DrawRectangle(new XPen(XColor.FromArgb(210, 210, 210), 0.6), new XSolidBrush(XColor.FromArgb(248, 250, 252)), bounds);
            gfx.DrawString("Summary", headerFont, XBrushes.Black,
                new XRect(bounds.X + 6, bounds.Y + 6, bounds.Width - 12, 14), XStringFormats.TopLeft);

            double y = bounds.Y + 24;
            for (int i = 0; i < chartPage.Summary.Items.Count; i++)
            {
                PdfSummaryItem item = chartPage.Summary.Items[i];
                gfx.DrawString(item.Metric, bodyFont, XBrushes.DimGray,
                    new XRect(bounds.X + 6, y, bounds.Width * 0.55, 12), XStringFormats.TopLeft);
                gfx.DrawString(item.Value, bodyFont, XBrushes.Black,
                    new XRect(bounds.X + bounds.Width * 0.5, y, bounds.Width * 0.48 - 6, 12), XStringFormats.TopRight);
                y += 12;
            }
        }

        private static void DrawHeader(XGraphics gfx, PdfPage page, PdfReport report)
        {
            PdfHeader header = report.Header;
            if (!HasHeaderContent(header))
            {
                return;
            }

            double left = Mm(report.Margins.Left);
            double right = Mm(report.Margins.Right);
            double top = Mm(report.Margins.Top);
            double width = page.Width - left - right;

            string fontFamily = !string.IsNullOrWhiteSpace(header.FontFamily) ? header.FontFamily! : PdfFontResolver.FamilyName;
            var titleFont = new XFont(fontFamily, header.FontSize, header.Bold ? XFontStyle.Bold : XFontStyle.Regular);
            double subFontSize = Math.Max(header.FontSize - 3.5, 7.5);
            var textFont = new XFont(fontFamily, subFontSize, XFontStyle.Regular);

            XBrush titleBrush = header.TitleColor != null
                ? new XSolidBrush(XColor.FromArgb(header.TitleColor.R, header.TitleColor.G, header.TitleColor.B))
                : XBrushes.Black;

            double x = left;
            if (header.Logo != null)
            {
                DrawImageAspectFit(gfx, header.Logo, new XRect(x, top, 52, 24));
                x += 58;
            }

            double textWidth = width - (x - left);
            bool hasTitle = !string.IsNullOrWhiteSpace(header.Title);
            bool hasText = !string.IsNullOrWhiteSpace(header.Text);

            if (hasTitle && hasText)
            {
                double titleHeight = Math.Ceiling(header.FontSize * 1.15);
                gfx.DrawString(header.Title!, titleFont, titleBrush,
                    new XRect(x, top, textWidth, titleHeight), XStringFormats.TopLeft);

                double textY = top + titleHeight + 2;
                double textHeight = Math.Ceiling(subFontSize * 1.2);
                gfx.DrawString(header.Text!, textFont, XBrushes.DimGray,
                    new XRect(x, textY, textWidth, textHeight), XStringFormats.TopLeft);
            }
            else if (hasTitle)
            {
                double titleHeight = Math.Ceiling(header.FontSize * 1.25);
                gfx.DrawString(header.Title!, titleFont, titleBrush,
                    new XRect(x, top + 4, textWidth, titleHeight), XStringFormats.TopLeft);
            }
            else if (hasText)
            {
                double textHeight = Math.Ceiling(subFontSize * 1.25);
                gfx.DrawString(header.Text!, textFont, XBrushes.DimGray,
                    new XRect(x, top + 4, textWidth, textHeight), XStringFormats.TopLeft);
            }

            // Divider line placed below subtitle, leaving ~8pt breathing room before content area
            double lineY = top + HeaderHeight - 8;
            gfx.DrawLine(new XPen(XColor.FromArgb(210, 210, 210), 0.6), left, lineY, left + width, lineY);
        }

        private static void DrawFooter(XGraphics gfx, PdfPage page, PdfReport report, int pageNo, int totalPages)
        {
            double left = Mm(report.Margins.Left);
            double right = Mm(report.Margins.Right);
            double y = page.Height - Mm(report.Margins.Bottom) - FooterHeight + 4;
            double width = page.Width - left - right;
            gfx.DrawLine(new XPen(XColor.FromArgb(200, 200, 200), 0.5), left, y, left + width, y);

            string footerFamily = !string.IsNullOrWhiteSpace(report.Footer.FontFamily) ? report.Footer.FontFamily! : PdfFontResolver.FamilyName;
            var font = new XFont(footerFamily, 8, XFontStyle.Regular);
            string numbers = string.Format(CultureInfo.InvariantCulture, report.Footer.PageNumberFormat ?? "{0} / {1}", pageNo, totalPages);
            gfx.DrawString(numbers, font, XBrushes.Gray,
                new XRect(left, y + 4, width, 14), XStringFormats.TopRight);

            double textX = left;
            if (report.Footer.Logo != null)
            {
                DrawImageAspectFit(gfx, report.Footer.Logo, new XRect(left, y + 3, 36, 14));
                textX += 42;
            }

            if (!string.IsNullOrWhiteSpace(report.Footer.Text))
            {
                gfx.DrawString(report.Footer.Text, font, XBrushes.Gray,
                    new XRect(textX, y + 4, width * 0.6, 14), XStringFormats.TopLeft);
            }
        }

        private static double DrawImageAspectFit(XGraphics gfx, PdfImage image, XRect box)
        {
            using (XImage xImage = XImage.FromStream(() => new MemoryStream(image.Bytes, writable: false)))
            {
                double scale = Math.Min(box.Width / xImage.PointWidth, box.Height / xImage.PointHeight);
                double w = xImage.PointWidth * scale;
                double h = xImage.PointHeight * scale;
                gfx.DrawImage(xImage, box.X, box.Y, w, h);
                return h;
            }
        }

        private static double Mm(double millimeters) => millimeters * 72.0 / 25.4;
    }
}
