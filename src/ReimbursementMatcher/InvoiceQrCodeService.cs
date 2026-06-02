using System.Globalization;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ZXing;
using ZXing.Common;

namespace ReimbursementMatcher;

public sealed class InvoiceQrMetadata
{
    public string Raw { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public decimal Amount { get; set; }
    public string InvoiceDate { get; set; } = "";
}

public static class InvoiceQrCodeService
{
    private static readonly BarcodeReaderGeneric Reader = new()
    {
        Options = new DecodingOptions
        {
            TryHarder = true,
            PossibleFormats = new List<BarcodeFormat> { BarcodeFormat.QR_CODE }
        }
    };

    private static readonly RGBLuminanceSource.BitmapFormat[] Formats =
    {
        RGBLuminanceSource.BitmapFormat.RGB24,
        RGBLuminanceSource.BitmapFormat.BGR24,
        RGBLuminanceSource.BitmapFormat.RGB32,
        RGBLuminanceSource.BitmapFormat.BGRA32,
        RGBLuminanceSource.BitmapFormat.Gray8
    };

    public static InvoiceQrMetadata ExtractFromPdfImages(string path)
    {
        if (!Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return new InvoiceQrMetadata();
        }

        try
        {
            using var document = UglyToad.PdfPig.PdfDocument.Open(path);
            foreach (var page in document.GetPages().Take(3))
            {
                foreach (var image in page.GetImages())
                {
                    if (!image.TryGetBytes(out var bytes))
                    {
                        var rawResult = DecodeRawImageStream(image.RawBytes.ToArray());
                        if (!string.IsNullOrWhiteSpace(rawResult))
                        {
                            var metadata = ParseInvoiceQr(rawResult);
                            if (!string.IsNullOrWhiteSpace(metadata.InvoiceNumber)
                                || metadata.Amount > 0
                                || !string.IsNullOrWhiteSpace(metadata.InvoiceDate))
                            {
                                return metadata;
                            }
                        }
                        continue;
                    }

                    var data = bytes.ToArray();
                    foreach (var format in Formats)
                    {
                        try
                        {
                            var result = Reader.Decode(data, image.WidthInSamples, image.HeightInSamples, format);
                            if (!string.IsNullOrWhiteSpace(result?.Text))
                            {
                                var metadata = ParseInvoiceQr(result.Text);
                                if (!string.IsNullOrWhiteSpace(metadata.InvoiceNumber)
                                    || metadata.Amount > 0
                                    || !string.IsNullOrWhiteSpace(metadata.InvoiceDate))
                                {
                                    return metadata;
                                }
                            }
                        }
                        catch
                        {
                            // Try the next pixel format.
                        }
                    }
                }
            }
        }
        catch
        {
            return new InvoiceQrMetadata();
        }

        return new InvoiceQrMetadata();
    }

    private static string DecodeRawImageStream(byte[] rawBytes)
    {
        try
        {
            using var stream = new MemoryStream(rawBytes);
            using var bitmap = new Bitmap(stream);
            return DecodeBitmap(bitmap);
        }
        catch
        {
            return "";
        }
    }

    private static string DecodeBitmap(Bitmap bitmap)
    {
        using var normalized = new Bitmap(bitmap.Width, bitmap.Height, PixelFormat.Format24bppRgb);
        using (var graphics = Graphics.FromImage(normalized))
        {
            graphics.DrawImage(bitmap, 0, 0, bitmap.Width, bitmap.Height);
        }

        var rect = new Rectangle(0, 0, normalized.Width, normalized.Height);
        var data = normalized.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        try
        {
            var stride = Math.Abs(data.Stride);
            var raw = new byte[stride * normalized.Height];
            Marshal.Copy(data.Scan0, raw, 0, raw.Length);

            var packed = new byte[normalized.Width * normalized.Height * 3];
            for (var y = 0; y < normalized.Height; y++)
            {
                Buffer.BlockCopy(raw, y * stride, packed, y * normalized.Width * 3, normalized.Width * 3);
            }

            var result = Reader.Decode(packed, normalized.Width, normalized.Height, RGBLuminanceSource.BitmapFormat.BGR24);
            return result?.Text ?? "";
        }
        finally
        {
            normalized.UnlockBits(data);
        }
    }

    private static InvoiceQrMetadata ParseInvoiceQr(string raw)
    {
        var parts = raw.Split(',').Select(p => p.Trim()).ToArray();
        var result = new InvoiceQrMetadata { Raw = raw };

        if (parts.Length >= 4)
        {
            var invoiceCode = NormalizeInvoiceNo(parts[2]);
            var invoiceNumber = NormalizeInvoiceNo(parts[3]);
            result.InvoiceNumber = invoiceCode.Length is >= 10 and <= 12 && invoiceNumber.Length == 8
                ? invoiceCode + invoiceNumber
                : invoiceNumber;
        }

        if (parts.Length >= 5
            && decimal.TryParse(parts[4], NumberStyles.Number, CultureInfo.InvariantCulture, out var amount)
            && amount > 0)
        {
            result.Amount = amount;
        }

        if (parts.Length >= 6)
        {
            result.InvoiceDate = NormalizeDate(parts[5]);
        }

        return result;
    }

    private static string NormalizeInvoiceNo(string value)
    {
        value = new string((value ?? "").Where(char.IsDigit).ToArray());
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static string NormalizeDate(string value)
    {
        value = new string((value ?? "").Where(char.IsDigit).ToArray());
        if (value.Length != 8)
        {
            return "";
        }

        var normalized = $"{value[..4]}-{value.Substring(4, 2)}-{value.Substring(6, 2)}";
        return DateTime.TryParse(normalized, out var parsed) ? parsed.ToString("yyyy-MM-dd") : "";
    }
}
