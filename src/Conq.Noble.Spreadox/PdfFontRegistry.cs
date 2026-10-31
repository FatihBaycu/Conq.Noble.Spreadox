using Conq.Noble.Spreadox.Internal;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Register custom TrueType fonts for PDF report rendering.
    /// Call <see cref="Register(string, byte[], bool, bool)"/> before generating any PDF.
    /// </summary>
    public static class PdfFontRegistry
    {
        /// <summary>
        /// Register a custom TrueType font from raw bytes.
        /// </summary>
        /// <param name="familyName">Font family name (e.g. "Roboto", "Open Sans").</param>
        /// <param name="ttfBytes">Raw TTF file bytes.</param>
        /// <param name="isBold">Whether this font file is the bold variant.</param>
        /// <param name="isItalic">Whether this font file is the italic variant.</param>
        public static void Register(string familyName, byte[] ttfBytes, bool isBold = false, bool isItalic = false)
        {
            PdfFontResolver.RegisterFont(familyName, ttfBytes, isBold, isItalic);
        }

        /// <summary>
        /// Register a custom TrueType font from a file path.
        /// </summary>
        /// <param name="familyName">Font family name (e.g. "Roboto", "Open Sans").</param>
        /// <param name="ttfFilePath">Path to the .ttf font file.</param>
        /// <param name="isBold">Whether this font file is the bold variant.</param>
        /// <param name="isItalic">Whether this font file is the italic variant.</param>
        public static void Register(string familyName, string ttfFilePath, bool isBold = false, bool isItalic = false)
        {
            PdfFontResolver.RegisterFont(familyName, ttfFilePath, isBold, isItalic);
        }

        /// <summary>
        /// Check whether a custom font family has been registered.
        /// </summary>
        public static bool IsRegistered(string familyName)
        {
            return PdfFontResolver.IsFamilyRegistered(familyName);
        }
    }
}
