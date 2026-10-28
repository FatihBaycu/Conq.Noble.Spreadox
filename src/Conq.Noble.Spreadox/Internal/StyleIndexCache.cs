using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Deduplicates fonts/fills/borders/numFmts/cellXfs for stylesheet emission.
    /// Hardened for large sheets: null/default styles short-circuit; date default is cached once.
    /// </summary>
    internal sealed class StyleIndexCache
    {
        private readonly Dictionary<string, uint> _numberFormatIds =
            new Dictionary<string, uint>(StringComparer.Ordinal);
        private readonly List<string> _customNumberFormats = new List<string>();

        private readonly Dictionary<string, uint> _fontIds = new Dictionary<string, uint>(StringComparer.Ordinal);
        private readonly List<ExcelFont> _fonts = new List<ExcelFont> { new ExcelFont() };

        private readonly Dictionary<string, uint> _fillIds = new Dictionary<string, uint>(StringComparer.Ordinal);
        private readonly List<ExcelFill> _fills = new List<ExcelFill>
        {
            new ExcelFill(), // none
            new ExcelFill { Pattern = ExcelFillPattern.None } // gray125 placeholder slot
        };

        private readonly Dictionary<string, uint> _borderIds = new Dictionary<string, uint>(StringComparer.Ordinal);
        private readonly List<ExcelBorders> _borders = new List<ExcelBorders> { new ExcelBorders() };

        private readonly Dictionary<string, uint> _cellFormatIds =
            new Dictionary<string, uint>(64, StringComparer.Ordinal);
        private readonly List<CellFormatSpec> _cellFormats = new List<CellFormatSpec>
        {
            new CellFormatSpec(0, 0, 0, 0, null, false, false, false, false, false)
        };

        private uint _nextCustomNumFmtId = 164;
        private uint? _defaultDateTimeStyleIndex;

        public StyleIndexCache()
        {
            // Seed fill index 1 as gray125 (required by Excel).
            _fillIds[FillKey(new ExcelFill())] = 0;
        }

        public int UniqueCellFormatCount => _cellFormats.Count;

        public int UniqueFontCount => _fonts.Count;

        public int UniqueFillCount => _fills.Count;

        public int UniqueBorderCount => _borders.Count;

        public int UniqueNumberFormatCount => _customNumberFormats.Count;

        public uint GetStyleIndex(ExcelCellStyle? style, bool applyDefaultDateFormat)
        {
            bool styleMissingOrDefault = style is null || style.IsDefault;
            if (styleMissingOrDefault && !applyDefaultDateFormat)
            {
                return 0;
            }

            if (styleMissingOrDefault && applyDefaultDateFormat)
            {
                return EnsureDefaultDateTimeStyle();
            }

            ExcelCellStyle effective = style!;
            string? numberFormat = effective.NumberFormat;
            if (string.IsNullOrEmpty(numberFormat) && applyDefaultDateFormat)
            {
                numberFormat = NumberFormats.DateTime;
            }

            uint numFmtId = ResolveNumberFormatId(numberFormat);
            uint fontId = ResolveFontId(effective.Font);
            uint fillId = ResolveFillId(effective.Fill);
            uint borderId = ResolveBorderId(effective.Borders);

            bool applyNumberFormat = !string.IsNullOrEmpty(numberFormat);
            bool applyFont = !effective.Font.IsDefault;
            bool applyFill = !effective.Fill.IsDefault;
            bool applyBorder = !effective.Borders.IsDefault;
            bool applyAlignment = !effective.Alignment.IsDefault;

            var spec = new CellFormatSpec(
                numFmtId,
                fontId,
                fillId,
                borderId,
                applyAlignment ? effective.Alignment.Clone() : null,
                applyNumberFormat,
                applyFont,
                applyFill,
                applyBorder,
                applyAlignment);

            string key = spec.Key;
            if (_cellFormatIds.TryGetValue(key, out uint index))
            {
                return index;
            }

            index = (uint)_cellFormats.Count;
            _cellFormats.Add(spec);
            _cellFormatIds[key] = index;
            return index;
        }

        private uint EnsureDefaultDateTimeStyle()
        {
            if (_defaultDateTimeStyleIndex.HasValue)
            {
                return _defaultDateTimeStyleIndex.Value;
            }

            var dateStyle = new ExcelCellStyle { NumberFormat = NumberFormats.DateTime };
            _defaultDateTimeStyleIndex = GetStyleIndex(dateStyle, applyDefaultDateFormat: false);
            return _defaultDateTimeStyleIndex.Value;
        }

        public Stylesheet BuildStylesheet()
        {
            var numberingFormats = new NumberingFormats();
            foreach (string code in _customNumberFormats)
            {
                uint id = _numberFormatIds[code];
                numberingFormats.AppendChild(new NumberingFormat
                {
                    NumberFormatId = id,
                    FormatCode = StringValue.FromString(code)
                });
            }

            numberingFormats.Count = (uint)_customNumberFormats.Count;

            var fonts = new Fonts();
            foreach (ExcelFont font in _fonts)
            {
                fonts.AppendChild(BuildFont(font));
            }

            fonts.Count = (uint)_fonts.Count;

            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));

            // indices 0 and 1 reserved; append custom fills from index 2
            for (int i = 2; i < _fills.Count; i++)
            {
                fills.AppendChild(BuildFill(_fills[i]));
            }

            fills.Count = (uint)(2 + Math.Max(0, _fills.Count - 2));

            var borders = new Borders();
            foreach (ExcelBorders border in _borders)
            {
                borders.AppendChild(BuildBorder(border));
            }

            borders.Count = (uint)_borders.Count;

            var cellStyleFormats = new CellStyleFormats(new CellFormat()) { Count = 1 };

            var cellFormats = new CellFormats();
            foreach (CellFormatSpec spec in _cellFormats)
            {
                var xf = new CellFormat
                {
                    NumberFormatId = spec.NumberFormatId,
                    FontId = spec.FontId,
                    FillId = spec.FillId,
                    BorderId = spec.BorderId,
                    FormatId = 0
                };

                if (spec.ApplyNumberFormat) xf.ApplyNumberFormat = true;
                if (spec.ApplyFont) xf.ApplyFont = true;
                if (spec.ApplyFill) xf.ApplyFill = true;
                if (spec.ApplyBorder) xf.ApplyBorder = true;
                if (spec.ApplyAlignment && spec.Alignment != null)
                {
                    xf.ApplyAlignment = true;
                    xf.Alignment = BuildAlignment(spec.Alignment);
                }

                cellFormats.AppendChild(xf);
            }

            cellFormats.Count = (uint)_cellFormats.Count;

            var stylesheet = new Stylesheet();
            if (_customNumberFormats.Count > 0)
            {
                stylesheet.AppendChild(numberingFormats);
            }

            stylesheet.AppendChild(fonts);
            stylesheet.AppendChild(fills);
            stylesheet.AppendChild(borders);
            stylesheet.AppendChild(cellStyleFormats);
            stylesheet.AppendChild(cellFormats);
            return stylesheet;
        }

        private uint ResolveNumberFormatId(string? formatCode)
        {
            if (string.IsNullOrWhiteSpace(formatCode) ||
                string.Equals(formatCode, NumberFormats.General, StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            formatCode = formatCode.Trim();
            if (TryBuiltIn(formatCode, out uint builtIn))
            {
                return builtIn;
            }

            if (_numberFormatIds.TryGetValue(formatCode, out uint id))
            {
                return id;
            }

            id = _nextCustomNumFmtId++;
            _numberFormatIds[formatCode] = id;
            _customNumberFormats.Add(formatCode);
            return id;
        }

        private static bool TryBuiltIn(string code, out uint id)
        {
            switch (code)
            {
                case "0": id = 1; return true;
                case "0.00": id = 2; return true;
                case "#,##0": id = 3; return true;
                case "#,##0.00": id = 4; return true;
                case "0%": id = 9; return true;
                case "0.00%": id = 10; return true;
                default:
                    id = 0;
                    return false;
            }
        }

        private uint ResolveFontId(ExcelFont font)
        {
            string key = FontKey(font);
            if (_fontIds.TryGetValue(key, out uint id))
            {
                return id;
            }

            // index 0 is default empty font already in list
            if (font.IsDefault)
            {
                _fontIds[key] = 0;
                return 0;
            }

            id = (uint)_fonts.Count;
            _fonts.Add(font.Clone());
            _fontIds[key] = id;
            return id;
        }

        private uint ResolveFillId(ExcelFill fill)
        {
            if (fill.IsDefault)
            {
                return 0;
            }

            string key = FillKey(fill);
            if (_fillIds.TryGetValue(key, out uint id))
            {
                return id;
            }

            // Ensure list has reserved slots 0,1 then custom from 2
            while (_fills.Count < 2)
            {
                _fills.Add(new ExcelFill());
            }

            id = (uint)_fills.Count;
            _fills.Add(fill.Clone());
            _fillIds[key] = id;
            return id;
        }

        private uint ResolveBorderId(ExcelBorders borders)
        {
            string key = BorderKey(borders);
            if (_borderIds.TryGetValue(key, out uint id))
            {
                return id;
            }

            if (borders.IsDefault)
            {
                _borderIds[key] = 0;
                return 0;
            }

            id = (uint)_borders.Count;
            _borders.Add(borders.Clone());
            _borderIds[key] = id;
            return id;
        }

        private static Font BuildFont(ExcelFont model)
        {
            var font = new Font();
            if (model.Bold) font.AppendChild(new Bold());
            if (model.Italic) font.AppendChild(new Italic());
            if (model.Underline) font.AppendChild(new Underline());
            if (model.Strikeout) font.AppendChild(new Strike());
            if (model.Size.HasValue)
            {
                font.AppendChild(new FontSize { Val = model.Size.Value });
            }

            if (model.Color != null)
            {
                font.AppendChild(new Color { Rgb = "FF" + model.Color.ToRgbHex() });
            }

            if (!string.IsNullOrEmpty(model.Name))
            {
                font.AppendChild(new FontName { Val = model.Name });
            }

            return font;
        }

        private static Fill BuildFill(ExcelFill model)
        {
            if (model.Pattern == ExcelFillPattern.Solid)
            {
                var pattern = new PatternFill { PatternType = PatternValues.Solid };
                ExcelColor? color = model.ForegroundColor ?? model.BackgroundColor;
                if (color != null)
                {
                    pattern.ForegroundColor = new ForegroundColor { Rgb = "FF" + color.ToRgbHex() };
                    pattern.BackgroundColor = new BackgroundColor { Rgb = "FF" + color.ToRgbHex() };
                }

                return new Fill(pattern);
            }

            return new Fill(new PatternFill { PatternType = PatternValues.None });
        }

        private static Border BuildBorder(ExcelBorders model)
        {
            return new Border(
                BuildBorderSide(model.Left, left: true),
                BuildBorderSide(model.Right, right: true),
                BuildBorderSide(model.Top, top: true),
                BuildBorderSide(model.Bottom, bottom: true),
                new DiagonalBorder());
        }

        private static OpenXmlElement BuildBorderSide(
            ExcelBorderEdge edge,
            bool left = false,
            bool right = false,
            bool top = false,
            bool bottom = false)
        {
            BorderStyleValues style = ToBorderStyle(edge.Style);
            OpenXmlElement element;
            if (left) element = new LeftBorder { Style = style };
            else if (right) element = new RightBorder { Style = style };
            else if (top) element = new TopBorder { Style = style };
            else element = new BottomBorder { Style = style };

            if (edge.Style != ExcelBorderStyle.None)
            {
                string rgb = edge.Color != null ? "FF" + edge.Color.ToRgbHex() : "FF000000";
                element.AppendChild(new Color { Rgb = rgb });
            }

            return element;
        }

        private static Alignment BuildAlignment(ExcelAlignment model)
        {
            var alignment = new Alignment
            {
                Horizontal = ToHorizontal(model.Horizontal),
                Vertical = ToVertical(model.Vertical),
                WrapText = model.WrapText
            };

            if (model.Indent > 0)
            {
                alignment.Indent = (uint)model.Indent;
            }

            return alignment;
        }

        private static BorderStyleValues ToBorderStyle(ExcelBorderStyle style)
        {
            switch (style)
            {
                case ExcelBorderStyle.Thin: return BorderStyleValues.Thin;
                case ExcelBorderStyle.Medium: return BorderStyleValues.Medium;
                case ExcelBorderStyle.Thick: return BorderStyleValues.Thick;
                case ExcelBorderStyle.Dashed: return BorderStyleValues.Dashed;
                case ExcelBorderStyle.Dotted: return BorderStyleValues.Dotted;
                case ExcelBorderStyle.Double: return BorderStyleValues.Double;
                case ExcelBorderStyle.Hair: return BorderStyleValues.Hair;
                case ExcelBorderStyle.MediumDashed: return BorderStyleValues.MediumDashed;
                case ExcelBorderStyle.DashDot: return BorderStyleValues.DashDot;
                case ExcelBorderStyle.MediumDashDot: return BorderStyleValues.MediumDashDot;
                case ExcelBorderStyle.DashDotDot: return BorderStyleValues.DashDotDot;
                case ExcelBorderStyle.MediumDashDotDot: return BorderStyleValues.MediumDashDotDot;
                case ExcelBorderStyle.SlantDashDot: return BorderStyleValues.SlantDashDot;
                default: return BorderStyleValues.None;
            }
        }

        private static HorizontalAlignmentValues ToHorizontal(ExcelHorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case ExcelHorizontalAlignment.Left: return HorizontalAlignmentValues.Left;
                case ExcelHorizontalAlignment.Center: return HorizontalAlignmentValues.Center;
                case ExcelHorizontalAlignment.Right: return HorizontalAlignmentValues.Right;
                case ExcelHorizontalAlignment.Fill: return HorizontalAlignmentValues.Fill;
                case ExcelHorizontalAlignment.Justify: return HorizontalAlignmentValues.Justify;
                case ExcelHorizontalAlignment.CenterContinuous: return HorizontalAlignmentValues.CenterContinuous;
                case ExcelHorizontalAlignment.Distributed: return HorizontalAlignmentValues.Distributed;
                default: return HorizontalAlignmentValues.General;
            }
        }

        private static VerticalAlignmentValues ToVertical(ExcelVerticalAlignment alignment)
        {
            switch (alignment)
            {
                case ExcelVerticalAlignment.Top: return VerticalAlignmentValues.Top;
                case ExcelVerticalAlignment.Center: return VerticalAlignmentValues.Center;
                case ExcelVerticalAlignment.Justify: return VerticalAlignmentValues.Justify;
                case ExcelVerticalAlignment.Distributed: return VerticalAlignmentValues.Distributed;
                default: return VerticalAlignmentValues.Bottom;
            }
        }

        private static string FontKey(ExcelFont font) =>
            string.Join("|",
                font.Name ?? "",
                font.Size?.ToString(CultureInfo.InvariantCulture) ?? "",
                font.Bold ? "1" : "0",
                font.Italic ? "1" : "0",
                font.Underline ? "1" : "0",
                font.Strikeout ? "1" : "0",
                font.Color?.ToRgbHex() ?? "");

        private static string FillKey(ExcelFill fill) =>
            string.Join("|",
                fill.Pattern.ToString(),
                fill.ForegroundColor?.ToRgbHex() ?? "",
                fill.BackgroundColor?.ToRgbHex() ?? "");

        private static string BorderKey(ExcelBorders b) =>
            string.Join("|",
                EdgeKey(b.Left), EdgeKey(b.Right), EdgeKey(b.Top), EdgeKey(b.Bottom));

        private static string EdgeKey(ExcelBorderEdge e) =>
            e.Style + ":" + (e.Color?.ToRgbHex() ?? "");

        private sealed class CellFormatSpec
        {
            public CellFormatSpec(
                uint numberFormatId,
                uint fontId,
                uint fillId,
                uint borderId,
                ExcelAlignment? alignment,
                bool applyNumberFormat,
                bool applyFont,
                bool applyFill,
                bool applyBorder,
                bool applyAlignment)
            {
                NumberFormatId = numberFormatId;
                FontId = fontId;
                FillId = fillId;
                BorderId = borderId;
                Alignment = alignment;
                ApplyNumberFormat = applyNumberFormat;
                ApplyFont = applyFont;
                ApplyFill = applyFill;
                ApplyBorder = applyBorder;
                ApplyAlignment = applyAlignment;
            }

            public uint NumberFormatId { get; }
            public uint FontId { get; }
            public uint FillId { get; }
            public uint BorderId { get; }
            public ExcelAlignment? Alignment { get; }
            public bool ApplyNumberFormat { get; }
            public bool ApplyFont { get; }
            public bool ApplyFill { get; }
            public bool ApplyBorder { get; }
            public bool ApplyAlignment { get; }

            public string Key =>
                string.Join("|",
                    NumberFormatId, FontId, FillId, BorderId,
                    ApplyNumberFormat, ApplyFont, ApplyFill, ApplyBorder, ApplyAlignment,
                    Alignment?.Horizontal, Alignment?.Vertical, Alignment?.WrapText, Alignment?.Indent);
        }
    }
}
