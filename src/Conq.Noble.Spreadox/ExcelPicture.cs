using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace Conq.Noble.Spreadox
{
    public enum ExcelPictureFormat
    {
        Png,
        Jpeg
    }

    /// <summary>
    /// Image anchored to a cell range (two-cell anchor). Size comes from the range, not pixel extents.
    /// </summary>
    public sealed class ExcelPicture
    {
        internal ExcelPicture(
            ExcelWorksheet worksheet,
            byte[] imageBytes,
            ExcelPictureFormat format,
            ExcelRange anchorRange,
            bool lockAspectRatio)
        {
            Worksheet = worksheet;
            ImageBytes = imageBytes;
            Format = format;
            AnchorRange = anchorRange;
            LockAspectRatio = lockAspectRatio;
        }

        public ExcelWorksheet Worksheet { get; }

        public byte[] ImageBytes { get; }

        public ExcelPictureFormat Format { get; }

        public ExcelRange AnchorRange { get; private set; }

        /// <summary>Maps to DrawingML <c>noChangeAspect</c>. Range still defines the box Excel places the picture in.</summary>
        public bool LockAspectRatio { get; set; }

        public string? Name { get; set; }

        public ExcelPicture SetAnchor(string a1Range)
        {
            AnchorRange = ExcelRange.Parse(a1Range);
            return this;
        }

        internal static ExcelPictureFormat DetectFormat(byte[] bytes)
        {
            if (bytes is null || bytes.Length < 8)
            {
                throw new ArgumentException("Image data is empty or too short.", nameof(bytes));
            }

            if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
            {
                return ExcelPictureFormat.Png;
            }

            if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            {
                return ExcelPictureFormat.Jpeg;
            }

            throw new ArgumentException("Only PNG and JPEG images are supported.", nameof(bytes));
        }

        internal static byte[] DecodeBase64(string data)
        {
            if (string.IsNullOrWhiteSpace(data))
            {
                throw new ArgumentException("Image data is required.", nameof(data));
            }

            string payload = data.Trim();
            if (payload.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                int comma = payload.IndexOf(',');
                if (comma < 0)
                {
                    throw new ArgumentException("Invalid data URI.", nameof(data));
                }

                payload = payload.Substring(comma + 1);
            }

            return Convert.FromBase64String(payload);
        }

        internal static byte[] Copy(byte[] bytes)
        {
            if (bytes is null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            var copy = new byte[bytes.Length];
            Buffer.BlockCopy(bytes, 0, copy, 0, bytes.Length);
            return copy;
        }
    }

    /// <summary>Pictures owned by a worksheet.</summary>
    public sealed class ExcelPictureCollection : IReadOnlyList<ExcelPicture>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly List<ExcelPicture> _items = new List<ExcelPicture>();

        internal ExcelPictureCollection(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        public int Count => _items.Count;

        public ExcelPicture this[int index]
        {
            get
            {
                _worksheet.Workbook.ThrowIfDisposed();
                return _items[index];
            }
        }

        public ExcelPicture Add(byte[] imageBytes, string a1Range, bool lockAspectRatio = true)
        {
            _worksheet.Workbook.ThrowIfDisposed();
            byte[] copy = ExcelPicture.Copy(imageBytes);
            ExcelPictureFormat format = ExcelPicture.DetectFormat(copy);
            var picture = new ExcelPicture(_worksheet, copy, format, ExcelRange.Parse(a1Range), lockAspectRatio);
            _items.Add(picture);
            return picture;
        }

        public ExcelPicture Add(Stream stream, string a1Range, bool lockAspectRatio = true)
        {
            if (stream is null) throw new ArgumentNullException(nameof(stream));
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                return Add(ms.ToArray(), a1Range, lockAspectRatio);
            }
        }

        public ExcelPicture AddFile(string path, string a1Range, bool lockAspectRatio = true)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path is required.", nameof(path));
            }

            return Add(File.ReadAllBytes(path), a1Range, lockAspectRatio);
        }

        /// <summary>Accepts raw base64 or a <c>data:image/png;base64,...</c> / JPEG data URI.</summary>
        public ExcelPicture AddFromBase64(string data, string a1Range, bool lockAspectRatio = true)
        {
            return Add(ExcelPicture.DecodeBase64(data), a1Range, lockAspectRatio);
        }

        public bool Remove(ExcelPicture picture)
        {
            _worksheet.Workbook.ThrowIfDisposed();
            if (picture is null) throw new ArgumentNullException(nameof(picture));
            return _items.Remove(picture);
        }

        public IEnumerator<ExcelPicture> GetEnumerator() => _items.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal void AddLoaded(ExcelPicture picture) => _items.Add(picture);
    }
}
