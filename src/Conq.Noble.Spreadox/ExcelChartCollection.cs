using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Charts owned by a worksheet.
    /// </summary>
    public sealed class ExcelChartCollection : IReadOnlyList<ExcelChart>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly List<ExcelChart> _charts = new List<ExcelChart>();

        internal ExcelChartCollection(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        public int Count => _charts.Count;

        public ExcelChart this[int index]
        {
            get
            {
                _worksheet.Workbook.ThrowIfDisposed();
                return _charts[index];
            }
        }

        /// <summary>Adds a native OOXML chart bound to this worksheet.</summary>
        public ExcelChart Add(ExcelChartType type, string? title = null)
        {
            _worksheet.Workbook.ThrowIfDisposed();
            if (!Enum.IsDefined(typeof(ExcelChartType), type))
            {
                throw new ArgumentOutOfRangeException(nameof(type));
            }

            var chart = new ExcelChart(_worksheet, type) { Title = title };
            _charts.Add(chart);
            return chart;
        }

        public bool Remove(ExcelChart chart)
        {
            _worksheet.Workbook.ThrowIfDisposed();
            if (chart is null) throw new ArgumentNullException(nameof(chart));
            return _charts.Remove(chart);
        }

        public void Clear() => _charts.Clear();

        public IEnumerator<ExcelChart> GetEnumerator() => _charts.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal void AddLoaded(ExcelChart chart) => _charts.Add(chart);

        internal IReadOnlyList<ExcelChart> AsReadOnly() =>
            new ReadOnlyCollection<ExcelChart>(_charts);
    }
}
