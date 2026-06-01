using System.Text;

namespace ReimbursementMatcher;

public static class TextEncodingFixer
{
    private static bool _registered;

    public static string Fix(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return value ?? "";
        }

        var text = value.Trim();
        if (!LooksLikeGbkMojibake(text))
        {
            return text;
        }

        RegisterCodePages();
        var repaired = DecodeSingleByteText(text, Encoding.GetEncoding("GB18030"));
        return Score(repaired) > Score(text) ? repaired : text;
    }

    private static void RegisterCodePages()
    {
        if (_registered) return;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _registered = true;
    }

    private static string DecodeSingleByteText(string text, Encoding targetEncoding)
    {
        var bytes = new List<byte>(text.Length);
        foreach (var ch in text)
        {
            if (ch <= 255)
            {
                bytes.Add((byte)ch);
            }
            else
            {
                return text;
            }
        }
        return targetEncoding.GetString(bytes.ToArray()).Trim();
    }

    private static bool LooksLikeGbkMojibake(string text)
    {
        if (text.Any(IsChinese))
        {
            return false;
        }

        var suspicious = text.Count(ch => ch is '¡' or '¿' or '¹' or 'ã' or '·' or '¢' or 'Ô' or 'Â' or '»' or 'î' or 'Ì' or 'á' or 'Ð' or 'Ñ' or 'Æ' or '±');
        var latin1High = text.Count(ch => ch is >= '\u00A0' and <= '\u00FF');
        return suspicious >= 2 || latin1High >= Math.Max(3, text.Length / 4);
    }

    private static int Score(string text)
    {
        var chinese = text.Count(IsChinese);
        var mojibake = text.Count(ch => ch is '¡' or '¿' or '¹' or 'ã' or '·' or '¢' or 'Ô' or 'Â' or '»' or 'î' or 'Ì' or 'á' or 'Ð' or 'Ñ');
        return chinese * 3 - mojibake * 2;
    }

    private static bool IsChinese(char ch)
    {
        return ch is >= '\u4e00' and <= '\u9fff';
    }
}
