using System;
using System.IO;
using Conq.Noble.Spreadox.Internal;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Streaming (SAX) workbook writer for large <c>sheetData</c>.
    /// Rows are flushed to the package incrementally and are not kept in an in-memory cell grid.
    /// Use <see cref="ExcelWorkbook"/> for random-access editing, charts, merges, and round-trip open/save.
    /// </summary>
    /// <remarks>
    /// Open XML SDK types are not part of this public surface; they remain an implementation detail.
    /// </remarks>
    public sealed class ExcelStreamingWorkbook : IDisposable
    {
        private readonly StreamingPackageWriter _writer;
        private bool _disposed;

        private ExcelStreamingWorkbook(StreamingPackageWriter writer)
        {
            _writer = writer;
        }

        public ExcelStreamingOptions Options
        {
            get
            {
                ThrowIfDisposed();
                return _writer.Options;
            }
        }

        /// <summary>Number of distinct cell formats registered so far (stylesheet dedupe).</summary>
        public int UniqueStyleCount
        {
            get
            {
                ThrowIfDisposed();
                return _writer.StyleCache.UniqueCellFormatCount;
            }
        }

        public static ExcelStreamingWorkbook Create(Stream stream, ExcelStreamingOptions? options = null)
        {
            if (stream is null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            return new ExcelStreamingWorkbook(StreamingPackageWriter.Create(stream, options ?? new ExcelStreamingOptions()));
        }

        public static ExcelStreamingWorkbook Create(string path, ExcelStreamingOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path is required.", nameof(path));
            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            FileStream stream = File.Create(path);
            try
            {
                return new ExcelStreamingWorkbook(
                    StreamingPackageWriter.Create(stream, options ?? new ExcelStreamingOptions(), ownsStream: true));
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Creates a style instance for reuse across many cells. Reuse the same instance so stylesheet
        /// dedupe collapses to one <c>cellXf</c>.
        /// </summary>
        public ExcelCellStyle CreateStyle()
        {
            ThrowIfDisposed();
            return new ExcelCellStyle();
        }

        public ExcelStreamingWorksheet AddWorksheet(string name)
        {
            ThrowIfDisposed();
            return _writer.AddWorksheet(name);
        }

        /// <summary>Finalizes styles, shared strings (if any), and closes the package.</summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _writer.Dispose();
            _disposed = true;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelStreamingWorkbook));
            }
        }
    }
}
