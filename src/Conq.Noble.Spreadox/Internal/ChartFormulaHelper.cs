using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Conq.Noble.Spreadox.Internal
{
    internal static class ChartFormulaHelper
    {
        public static string ToSheetLocalFormula(string sheetName, string a1Range)
        {
            ExcelRange range = ExcelRange.Parse(a1Range);
            return ToSheetLocalFormula(sheetName, range);
        }

        public static string ToSheetLocalFormula(string sheetName, ExcelRange range)
        {
            string escaped = EscapeSheetName(sheetName);
            string start = AbsoluteA1(range.StartRow, range.StartColumn);
            if (range.StartRow == range.EndRow && range.StartColumn == range.EndColumn)
            {
                return escaped + "!" + start;
            }

            string end = AbsoluteA1(range.EndRow, range.EndColumn);
            return escaped + "!" + start + ":" + end;
        }

        public static string EscapeSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                return "Sheet1";
            }

            bool needsQuotes = false;
            foreach (char ch in sheetName)
            {
                if (!(char.IsLetterOrDigit(ch) || ch == '_'))
                {
                    needsQuotes = true;
                    break;
                }
            }

            if (!needsQuotes)
            {
                return sheetName;
            }

            return "'" + sheetName.Replace("'", "''") + "'";
        }

        public static string AbsoluteA1(int row, int column)
        {
            return "$" + ColumnLetters(column) + "$" + row.ToString(CultureInfo.InvariantCulture);
        }

        public static string ColumnLetters(int column)
        {
            var sb = new StringBuilder();
            int c = column;
            while (c > 0)
            {
                c--;
                sb.Insert(0, (char)('A' + (c % 26)));
                c /= 26;
            }

            return sb.ToString();
        }

        public static List<object?> ReadLinearValues(ExcelWorksheet sheet, string a1Range)
        {
            ExcelRange range = ExcelRange.Parse(a1Range);
            var values = new List<object?>();

            bool singleColumn = range.StartColumn == range.EndColumn;
            bool singleRow = range.StartRow == range.EndRow;

            if (singleColumn)
            {
                for (int r = range.StartRow; r <= range.EndRow; r++)
                {
                    values.Add(GetCellObject(sheet, r, range.StartColumn));
                }
            }
            else if (singleRow)
            {
                for (int c = range.StartColumn; c <= range.EndColumn; c++)
                {
                    values.Add(GetCellObject(sheet, range.StartRow, c));
                }
            }
            else
            {
                // Flatten row-major for unusual 2D ranges.
                for (int r = range.StartRow; r <= range.EndRow; r++)
                {
                    for (int c = range.StartColumn; c <= range.EndColumn; c++)
                    {
                        values.Add(GetCellObject(sheet, r, c));
                    }
                }
            }

            return values;
        }

        public static string? ToCacheString(object? value)
        {
            if (value is null)
            {
                return null;
            }

            switch (value)
            {
                case string s:
                    return s;
                case DateTime dt:
                    return dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                case bool b:
                    return b ? "TRUE" : "FALSE";
                case IFormattable f:
                    return f.ToString(null, CultureInfo.InvariantCulture);
                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        public static double? ToCacheNumber(object? value)
        {
            if (value is null)
            {
                return null;
            }

            switch (value)
            {
                case double d:
                    return d;
                case float f:
                    return f;
                case int i:
                    return i;
                case long l:
                    return l;
                case decimal m:
                    return (double)m;
                case bool b:
                    return b ? 1d : 0d;
                case DateTime dt:
                    return dt.ToOADate();
                case string s:
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
                    {
                        return parsed;
                    }

                    return null;
                default:
                    try
                    {
                        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        return null;
                    }
            }
        }

        private static object? GetCellObject(ExcelWorksheet sheet, int row, int column)
        {
            if (!sheet.TryGetCell(row, column, out ExcelCell? cell) || cell is null || cell.IsEmpty)
            {
                return null;
            }

            if (cell.HasFormula && cell.Value != null)
            {
                return cell.Value;
            }

            return cell.Value;
        }
    }
}
