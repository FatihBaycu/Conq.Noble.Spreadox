using System;
using System.IO;

namespace Conq.Noble.Spreadox
{
    /// <summary>PNG or JPEG image for PDF logos (aspect-fit when drawn).</summary>
    public sealed class PdfImage
    {
        private PdfImage(byte[] bytes, ExcelPictureFormat format)
        {
            Bytes = bytes;
            Format = format;
        }

        public byte[] Bytes { get; }

        public ExcelPictureFormat Format { get; }

        public static PdfImage FromBytes(byte[] bytes)
        {
            if (bytes is null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            byte[] copy = (byte[])bytes.Clone();
            return new PdfImage(copy, ExcelPicture.DetectFormat(copy));
        }

        public static PdfImage FromFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path is required.", nameof(path));
            }

            return FromBytes(File.ReadAllBytes(path));
        }
    }
}
