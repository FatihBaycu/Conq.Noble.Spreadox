using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using Conq.Noble.Spreadox.Internal;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// A4 PDF report: cover page plus one page per chart. Charts are drawn as vectors,
    /// not OOXML. Independent of <see cref="ExcelWorkbook"/>.
    /// </summary>
    public sealed class PdfReport : IDisposable
    {
        private readonly List<PdfChartPage> _pages = new List<PdfChartPage>();
        private bool _disposed;

        private PdfReport(PdfPageSize pageSize)
        {
            PageSize = pageSize;
        }

        public PdfPageSize PageSize { get; }

        public PdfMargins Margins { get; set; } = PdfMargins.Default;

        public PdfCoverPage Cover { get; } = new PdfCoverPage();

        public PdfHeader Header { get; } = new PdfHeader();

        public PdfFooter Footer { get; } = new PdfFooter();

        public IReadOnlyList<PdfChartPage> ChartPages => new ReadOnlyCollection<PdfChartPage>(_pages);

        public static PdfReport Create(PdfPageSize pageSize = PdfPageSize.A4Landscape)
        {
            if (pageSize != PdfPageSize.A4Landscape)
            {
                throw new ArgumentOutOfRangeException(nameof(pageSize), "Only A4 landscape is supported.");
            }

            return new PdfReport(pageSize);
        }

        public PdfChartPage AddChartPage(string title)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(title))
            {
                throw new ArgumentException("Title is required.", nameof(title));
            }

            var page = new PdfChartPage(title);
            _pages.Add(page);
            return page;
        }

        public byte[] ToByteArray()
        {
            ThrowIfDisposed();
            using (var stream = new MemoryStream())
            {
                Save(stream);
                return stream.ToArray();
            }
        }

        public void Save(string path)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path is required.", nameof(path));
            }

            using (FileStream stream = File.Create(path))
            {
                Save(stream);
            }
        }

        public void Save(Stream stream)
        {
            ThrowIfDisposed();
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            PdfReportWriter.Write(this, stream);
        }

        public void Dispose()
        {
            _disposed = true;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(PdfReport));
            }
        }
    }
}
