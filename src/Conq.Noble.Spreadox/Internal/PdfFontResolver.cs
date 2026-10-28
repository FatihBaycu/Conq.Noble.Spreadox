using System;
using System.IO;
using System.Reflection;
using PdfSharpCore.Fonts;

namespace Conq.Noble.Spreadox.Internal
{
    internal sealed class PdfFontResolver : IFontResolver
    {
        internal const string FamilyName = "Liberation Sans";

        private const string FaceRegular = "LiberationSans#R";
        private const string FaceBold = "LiberationSans#B";
        private const string FaceItalic = "LiberationSans#I";
        private const string FaceBoldItalic = "LiberationSans#BI";

        private static readonly object Gate = new object();
        private static bool _installed;

        public string DefaultFontName => FamilyName;

        internal static void EnsureInstalled()
        {
            if (_installed)
            {
                return;
            }

            lock (Gate)
            {
                if (_installed)
                {
                    return;
                }

                GlobalFontSettings.FontResolver = new PdfFontResolver();
                _installed = true;
            }
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (isBold && isItalic)
            {
                return new FontResolverInfo(FaceBoldItalic);
            }

            if (isBold)
            {
                return new FontResolverInfo(FaceBold);
            }

            if (isItalic)
            {
                return new FontResolverInfo(FaceItalic);
            }

            return new FontResolverInfo(FaceRegular);
        }

        public byte[] GetFont(string faceName)
        {
            string resource = faceName switch
            {
                FaceBold => "LiberationSans-Bold.ttf",
                FaceItalic => "LiberationSans-Italic.ttf",
                FaceBoldItalic => "LiberationSans-BoldItalic.ttf",
                _ => "LiberationSans-Regular.ttf"
            };

            return LoadEmbedded(resource);
        }

        private static byte[] LoadEmbedded(string fileName)
        {
            Assembly assembly = typeof(PdfFontResolver).Assembly;
            string resourceName = "Conq.Noble.Spreadox.Fonts." + fileName;
            Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream is null)
            {
                throw new InvalidOperationException("Embedded font not found: " + resourceName);
            }

            using (stream)
            using (var copy = new MemoryStream())
            {
                stream.CopyTo(copy);
                return copy.ToArray();
            }
        }
    }
}
