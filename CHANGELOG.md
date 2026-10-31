# Changelog

All notable changes to the **Spreadox** library will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.9.0] - 2026-10-31

### Added
- **`PdfHeader` Support:** Added full running header customization on `PdfReport` with `Title`, `Text` (subtitle), `Logo` (`PdfImage`), `FontFamily`, `FontSize`, `Bold`, and `TitleColor` (`ExcelColor`).
- **Runtime TrueType (`.ttf`) Font Registration (`PdfFontRegistry`):** Added public API to dynamically register custom TrueType fonts at runtime via `PdfFontRegistry.Register(familyName, ttfBytes)` or file path.
- **Font Support in `PdfFooter`:** Added `FontFamily` property to `PdfFooter` to allow using registered custom fonts in running footers.
- **Report Studio MVC Sample (`samples/ReportBuilderMvc`):**
  - Interactive ASP.NET Core MVC web application with live chart builder.
  - Dynamic random data generator with configurable series count and inline data editing.
  - Per-series chart type picker (`ColumnClustered`, `BarClustered`, `Line`, `LineMarker`, `Area`, `Pie`, `Doughnut`, `Scatter`, `Radar`, `Combo`).
  - Runtime `.ttf` font file upload & live registration.
  - Single-page vs. separate-page layout options for both Excel (`.xlsx`) and PDF (`.pdf`) exports.
  - Scalable data-driven multi-language support (English default, Turkish) with a floating glassmorphic language switcher.
- **Visual Documentation:** Added high-resolution (2x Retina) vector PDF and Excel output screenshots under `docs/images/`.

### Changed
- **`PdfReportWriter` Spacing:** Increased `HeaderHeight` to 36pt, refined dynamic subtitle vertical offset, and added an 8pt (~3mm) breathing margin below the header divider line so it never overlaps subtitle text.
- **Documentation:** Updated `README.md` with visual preview galleries and custom font code walkthroughs.

---

## [0.8.0] - 2026-10-15

### Added
- **Excel DOM Workbook (`ExcelWorkbook`):**
  - Microsoft Open XML SDK based spreadsheet engine free of OpenXML public types.
  - Native OOXML DrawingML chart support (Column, Bar, Line, Scatter, Area, Pie, Doughnut, Radar, Bubble, Combo, BoxAndWhisker).
  - Cell formatting: Font, Fill, Border styles, Alignment, Number Formats.
  - Conditional formatting (Comparison rules and 3-color scales).
  - Merged cells, Freeze panes, and Auto-filter support.
- **High-Volume Streaming Engine (`ExcelStreamingWorkbook`):** Low-memory forward-only writer capable of flushing 100,000+ rows directly to disk.
- **Vector PDF Engine (`PdfReport`):**
  - A4 Landscape publication-ready PDF chart reports.
  - Vector-drawn charts from in-memory arrays (not OOXML charts).
  - Cover page support with metadata facts, company title, and date.
  - Running footer with page numbering format (`Page {0} / {1}`) and logo.
  - Right-hand metric summary sidebar (up to 28 rows).
  - Embedded SIL Open Font License Liberation Sans fonts.

