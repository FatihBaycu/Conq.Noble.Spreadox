using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using Conq.Noble.Spreadox.Internal;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// High-level workbook. Public API is independent of Open XML SDK types.
    /// Serialize/deserialize via <see cref="Save(string)"/> / <see cref="Open(string)"/>.
    /// </summary>
    public sealed class ExcelWorkbook : IDisposable
    {
        private readonly List<ExcelWorksheet> _worksheets = new List<ExcelWorksheet>();
        private readonly List<ExcelNamedRange> _namedRanges = new List<ExcelNamedRange>();
        private bool _disposed;

        private ExcelWorkbook()
        {
        }

        /// <summary>
        /// Creates an empty workbook with a single sheet named "Sheet1".
        /// </summary>
        public static ExcelWorkbook Create()
        {
            var workbook = new ExcelWorkbook();
            workbook.AddWorksheet("Sheet1");
            return workbook;
        }

        public static ExcelWorkbook Open(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path is required.", nameof(path));
            }

            using (FileStream stream = File.OpenRead(path))
            {
                return Open(stream);
            }
        }

        public static ExcelWorkbook Open(Stream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (!stream.CanRead)
            {
                throw new ArgumentException("Stream must be readable.", nameof(stream));
            }

            var workbook = new ExcelWorkbook();
            OpenXmlWorkbookReader.Load(workbook, stream);
            if (workbook._worksheets.Count == 0)
            {
                workbook.AddWorksheet("Sheet1");
            }

            return workbook;
        }

        /// <summary>Worksheets in this workbook (at least one while the workbook is alive).</summary>
        public IReadOnlyList<ExcelWorksheet> Worksheets
        {
            get
            {
                ThrowIfDisposed();
                return new ReadOnlyCollection<ExcelWorksheet>(_worksheets);
            }
        }

        /// <summary>Defined names owned by the workbook.</summary>
        public IReadOnlyList<ExcelNamedRange> NamedRanges
        {
            get
            {
                ThrowIfDisposed();
                return new ReadOnlyCollection<ExcelNamedRange>(_namedRanges);
            }
        }

        public ExcelNamedRange AddNamedRange(string name, string refersTo)
        {
            ThrowIfDisposed();
            if (name is null) throw new ArgumentNullException(nameof(name));
            if (refersTo is null) throw new ArgumentNullException(nameof(refersTo));

            foreach (ExcelNamedRange existing in _namedRanges)
            {
                if (string.Equals(existing.Name, name.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"A named range '{name.Trim()}' already exists.");
                }
            }

            var named = new ExcelNamedRange(name, refersTo);
            _namedRanges.Add(named);
            return named;
        }

        public bool RemoveNamedRange(string name)
        {
            ThrowIfDisposed();
            for (int i = 0; i < _namedRanges.Count; i++)
            {
                if (string.Equals(_namedRanges[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    _namedRanges.RemoveAt(i);
                    return true;
                }
            }

            return false;
        }

        public ExcelWorksheet this[string name]
        {
            get
            {
                ThrowIfDisposed();
                if (name is null) throw new ArgumentNullException(nameof(name));

                ExcelWorksheet? sheet = FindWorksheet(name);
                if (sheet is null)
                {
                    throw new KeyNotFoundException($"Worksheet '{name}' was not found.");
                }

                return sheet;
            }
        }

        public ExcelWorksheet AddWorksheet(string name)
        {
            ThrowIfDisposed();
            if (name is null) throw new ArgumentNullException(nameof(name));
            EnsureWorksheetNameAvailable(name, exclude: null);
            var sheet = new ExcelWorksheet(this, name);
            _worksheets.Add(sheet);
            return sheet;
        }

        public bool RemoveWorksheet(string name)
        {
            ThrowIfDisposed();
            ExcelWorksheet? sheet = FindWorksheet(name);
            if (sheet is null)
            {
                return false;
            }

            return RemoveWorksheet(sheet);
        }

        public bool RemoveWorksheet(ExcelWorksheet worksheet)
        {
            ThrowIfDisposed();
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (!ReferenceEquals(worksheet.Workbook, this))
            {
                throw new ArgumentException("Worksheet does not belong to this workbook.", nameof(worksheet));
            }

            if (_worksheets.Count <= 1)
            {
                throw new InvalidOperationException("A workbook must contain at least one worksheet.");
            }

            return _worksheets.Remove(worksheet);
        }

        public void RenameWorksheet(string currentName, string newName)
        {
            ThrowIfDisposed();
            this[currentName].Name = newName;
        }

        public void Save(string path)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path is required.", nameof(path));
            }

            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
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

            if (!stream.CanWrite)
            {
                throw new ArgumentException("Stream must be writable.", nameof(stream));
            }

            OpenXmlWorkbookWriter.Write(this, stream);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _worksheets.Clear();
            _namedRanges.Clear();
            _disposed = true;
        }

        internal void ClearNamedRangesForLoad() => _namedRanges.Clear();

        internal void AddLoadedNamedRange(ExcelNamedRange namedRange) => _namedRanges.Add(namedRange);

        internal void EnsureWorksheetNameAvailable(string name, ExcelWorksheet? exclude)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Worksheet name is required.", nameof(name));
            }

            foreach (ExcelWorksheet sheet in _worksheets)
            {
                if (exclude != null && ReferenceEquals(sheet, exclude))
                {
                    continue;
                }

                if (string.Equals(sheet.Name, name.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"A worksheet named '{name.Trim()}' already exists.");
                }
            }
        }

        internal void AddLoadedWorksheet(ExcelWorksheet worksheet)
        {
            _worksheets.Add(worksheet);
        }

        internal void ClearWorksheetsForLoad()
        {
            _worksheets.Clear();
        }

        private ExcelWorksheet? FindWorksheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            foreach (ExcelWorksheet sheet in _worksheets)
            {
                if (string.Equals(sheet.Name, name.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        internal void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelWorkbook));
            }
        }
    }
}
