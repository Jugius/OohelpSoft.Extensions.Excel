using OfficeOpenXml;
using System.Drawing;
using System.Globalization;

namespace OohelpSoft.Excel.Utils;
public static class BackgroundColor
{
    public static void SetBackgroundColor(this ExcelRangeBase range, Color color)
    {
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
        string colorcode = hexColor.TrimStart('#');

        if (colorcode == string.Empty)
            return Color.Empty;

        if (colorcode.Length == 6)
            return Color.FromArgb(255, // hardcoded opaque
                        int.Parse(colorcode[..2], NumberStyles.HexNumber),
                        int.Parse(colorcode.Substring(2, 2), NumberStyles.HexNumber),
                        int.Parse(colorcode.Substring(4, 2), NumberStyles.HexNumber));
        else if (colorcode.Length == 8) // assuming length of 8
            return Color.FromArgb(
                        int.Parse(colorcode[..2], NumberStyles.HexNumber),
                        int.Parse(colorcode.Substring(2, 2), NumberStyles.HexNumber),
                        int.Parse(colorcode.Substring(4, 2), NumberStyles.HexNumber),
                        int.Parse(colorcode.Substring(6, 2), NumberStyles.HexNumber));
        else
            throw new Exception($"'{colorcode}' is not valid Color");
    }
}
