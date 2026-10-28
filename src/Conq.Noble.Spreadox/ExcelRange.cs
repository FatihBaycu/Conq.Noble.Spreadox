using System;
using System.Globalization;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Inclusive rectangular range in 1-based row/column coordinates.
    /// </summary>
    public readonly struct ExcelRange : IEquatable<ExcelRange>
    {
        public ExcelRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            if (startRow <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startColumn));
            if (endRow < startRow) throw new ArgumentOutOfRangeException(nameof(endRow));
            if (endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(endColumn));

            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }

        public int StartRow { get; }
        public int StartColumn { get; }
        public int EndRow { get; }
        public int EndColumn { get; }

        public string A1 =>
            StartRow == EndRow && StartColumn == EndColumn
                ? CellAddress.Format(StartRow, StartColumn)
                : CellAddress.Format(StartRow, StartColumn) + ":" + CellAddress.Format(EndRow, EndColumn);

        public static ExcelRange Parse(string a1Range)
        {
            if (string.IsNullOrWhiteSpace(a1Range))
            {
                throw new ArgumentException("Range is required.", nameof(a1Range));
            }

            string text = a1Range.Trim().ToUpperInvariant();
            int colon = text.IndexOf(':');
            if (colon < 0)
            {
                CellAddress.Parse(text, out int row, out int col);
                return new ExcelRange(row, col, row, col);
            }

            CellAddress.Parse(text.Substring(0, colon), out int r1, out int c1);
            CellAddress.Parse(text.Substring(colon + 1), out int r2, out int c2);
            return new ExcelRange(
                Math.Min(r1, r2),
                Math.Min(c1, c2),
                Math.Max(r1, r2),
                Math.Max(c1, c2));
        }

        public bool Equals(ExcelRange other) =>
            StartRow == other.StartRow
            && StartColumn == other.StartColumn
            && EndRow == other.EndRow
            && EndColumn == other.EndColumn;

        public override bool Equals(object? obj) => obj is ExcelRange other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = StartRow;
                hash = (hash * 397) ^ StartColumn;
                hash = (hash * 397) ^ EndRow;
                hash = (hash * 397) ^ EndColumn;
                return hash;
            }
        }

        public override string ToString() => A1;
    }

    /// <summary>
    /// Workbook-level defined name.
    /// </summary>
    public sealed class ExcelNamedRange
    {
        public ExcelNamedRange(string name, string refersTo)
        {
            Name = ValidateName(name);
            RefersTo = ValidateRefersTo(refersTo);
        }

        public string Name { get; }
        public string RefersTo { get; set; }

        private static string ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Named range name is required.", nameof(name));
            }

            name = name.Trim();
            if (name.Length > 255)
            {
                throw new ArgumentException("Named range name is too long.", nameof(name));
            }

            return name;
        }

        private static string ValidateRefersTo(string refersTo)
        {
            if (string.IsNullOrWhiteSpace(refersTo))
            {
                throw new ArgumentException("Named range reference is required.", nameof(refersTo));
            }

            return refersTo.Trim();
        }
    }

    public enum ExcelConditionalFormatKind
    {
        CellValue,
        ColorScale
    }

    /// <summary>One stop on a 2- or 3-color scale.</summary>
    public sealed class ExcelColorScaleStop
    {
        public ExcelColorScaleStop(ExcelColorScaleValueType type, double? value, ExcelColor color)
        {
            Type = type;
            Value = value;
            Color = color ?? throw new ArgumentNullException(nameof(color));
        }

        public ExcelColorScaleValueType Type { get; }
        public double? Value { get; }
        public ExcelColor Color { get; }
    }

    public enum ExcelColorScaleValueType
    {
        Percentile,
        Percent,
        Min,
        Max,
        Number
    }

    /// <summary>3-color (or 2-color) color scale.</summary>
    public sealed class ExcelColorScale
    {
        public ExcelColorScale(ExcelColorScaleStop low, ExcelColorScaleStop high, ExcelColorScaleStop? midpoint = null)
        {
            Low = low ?? throw new ArgumentNullException(nameof(low));
            High = high ?? throw new ArgumentNullException(nameof(high));
            Midpoint = midpoint;
        }

        public ExcelColorScaleStop Low { get; }
        public ExcelColorScaleStop? Midpoint { get; }
        public ExcelColorScaleStop High { get; }
    }

    /// <summary>
    /// Conditional formatting rule: cell-value comparison or color scale.
    /// </summary>
    public sealed class ExcelConditionalFormat
    {
        public ExcelConditionalFormat(
            string rangeA1,
            ConditionalOperator op,
            string formulaOperand,
            ExcelColor? fillColor = null)
        {
            Range = ExcelRange.Parse(rangeA1);
            Kind = ExcelConditionalFormatKind.CellValue;
            Operator = op;
            FormulaOperand = formulaOperand ?? throw new ArgumentNullException(nameof(formulaOperand));
            FillColor = fillColor ?? ExcelColor.FromRgb(255, 199, 206);
        }

        private ExcelConditionalFormat(string rangeA1, ExcelColorScale scale)
        {
            Range = ExcelRange.Parse(rangeA1);
            Kind = ExcelConditionalFormatKind.ColorScale;
            ColorScale = scale ?? throw new ArgumentNullException(nameof(scale));
            Operator = ConditionalOperator.GreaterThan;
            FormulaOperand = string.Empty;
            FillColor = scale.Low.Color;
        }

        public static ExcelConditionalFormat ColorScalePercentile(
            string rangeA1,
            ExcelColor low,
            ExcelColor midpoint,
            ExcelColor high)
        {
            var scale = new ExcelColorScale(
                new ExcelColorScaleStop(ExcelColorScaleValueType.Percentile, 0, low),
                new ExcelColorScaleStop(ExcelColorScaleValueType.Percentile, 100, high),
                new ExcelColorScaleStop(ExcelColorScaleValueType.Percentile, 50, midpoint));
            return new ExcelConditionalFormat(rangeA1, scale);
        }

        public ExcelRange Range { get; }
        public ExcelConditionalFormatKind Kind { get; }
        public ConditionalOperator Operator { get; }
        public string FormulaOperand { get; }
        public ExcelColor FillColor { get; }
        public ExcelColorScale? ColorScale { get; }
        public int Priority { get; set; } = 1;
    }
}
