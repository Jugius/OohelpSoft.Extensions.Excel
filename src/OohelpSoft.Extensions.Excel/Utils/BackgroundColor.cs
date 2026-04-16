using OfficeOpenXml;
using System.Drawing;
using System.Globalization;

namespace OohelpSoft.Excel.Utils;
public static class BackgroundColor
{
    public static void SetBackgroundColor(this ExcelRangeBase range, Color color)
    {
        ArgumentNullException.ThrowIfNull(range);

        if (color == Color.Empty)
        {
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
        }
        else
        {
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }
    }
    public static Color GetBackgroundColor(this ExcelRangeBase cell)
    {
        ArgumentNullException.ThrowIfNull(cell);

        var fill = cell.Style.Fill;

        if (fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.None)
            return Color.Empty;

        string colorHex;
        if (fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.Solid)
        {
            colorHex = !string.IsNullOrEmpty(fill.BackgroundColor.Rgb)
                ? fill.BackgroundColor.Rgb
                : fill.BackgroundColor.LookupColor();
        }
        else
        {
            colorHex = !string.IsNullOrEmpty(fill.PatternColor.Rgb)
                ? fill.PatternColor.Rgb
                : fill.PatternColor.LookupColor();
        }

        return ConvertColor(colorHex);
    }
    private static Color ConvertColor(string hexColor)
    {
        if (string.IsNullOrWhiteSpace(hexColor))
            return Color.Empty;

        ReadOnlySpan<char> span = hexColor.AsSpan().TrimStart('#');

        if (span.IsEmpty)
            return Color.Empty;

        if (span.Length != 6 && span.Length != 8)
            throw new ArgumentException($"'{hexColor}' is not a valid color");

        byte a = 255;
        int offset = 0;

        if (span.Length == 8)
        {
            a = ParseByte(span, 0);
            offset = 2;
        }

        byte r = ParseByte(span, offset);
        byte g = ParseByte(span, offset + 2);
        byte b = ParseByte(span, offset + 4);

        return Color.FromArgb(a, r, g, b);
    }
    private static byte ParseByte(ReadOnlySpan<char> span, int start)
    {
        return byte.Parse(span.Slice(start, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
    }
}
