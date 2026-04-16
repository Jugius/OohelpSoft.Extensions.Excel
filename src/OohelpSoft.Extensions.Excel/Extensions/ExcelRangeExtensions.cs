using OfficeOpenXml;
using OfficeOpenXml.Style;
using OohelpSoft.Excel.DocumentModel.Styles;
using OohelpSoft.Excel.Utils;

namespace OohelpSoft.Excel.Extensions;
internal static class ExcelRangeExtensions
{
    public static void ApplyFont(this ExcelRangeBase range, DocumentModel.ExcelFont font)
    {
        range.Style.Font.Name = font.Name;
        range.Style.Font.Size = font.Size;
        range.Style.Font.Bold = font.Bold;
        range.Style.Font.Italic = font.Italic;
        if (font.Color.HasValue)
            range.Style.Font.Color.SetColor(font.Color.Value);
    }

    public static void ApplyTableStyles(this ExcelRangeBase range, ExcelTableStyle style)
    {
        using var header = range.Offset(0, 0, 1, range.Columns);
        header.ApplyHeaderStyle(style.Header);

        using var body = range.Offset(1, 0, range.Rows - 1, range.Columns);
        body.ApplyBodyStyle(style);

        if (style.Header.Color.HasValue)
        { 
            using var lastRow = range.Offset(range.Rows - 1, 0, 1, range.Columns);
            lastRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            lastRow.Style.Border.Bottom.Color.SetColor(style.Header.Color.Value);

            using var lastCol = range.Offset(0, range.Columns - 1, range.Rows, 1);
            lastCol.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            lastCol.Style.Border.Right.Color.SetColor(style.Header.Color.Value);
        }
    }
    private static void ApplyHeaderStyle(this ExcelRangeBase range, ExcelTableHeaderStyle style)
    {
        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        range.Style.WrapText = style.WrapText;
        
        if (style.Height.HasValue)
            range.EntireRow.Height = style.Height.Value;

        range.ApplyFont(style.Font);

        if (style.Color.HasValue)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Color.SetColor(style.Color.Value);
            range.SetBackgroundColor(style.Color.Value);
        }        

        if (style.Font.Color.HasValue)
        {
            var headercolor = style.Font.Color.Value;

            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;            
            range.Style.Border.Left.Color.SetColor(headercolor);

            if (style.Color.HasValue)
            {
                using var firstCol = range.Offset(0, 0, range.Rows, 1);
                firstCol.Style.Border.Left.Color.SetColor(style.Color.Value);
            }            
        }
    }
    private static void ApplyBodyStyle(this ExcelRangeBase range, ExcelTableStyle style)
    {
        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        range.ApplyFont(style.Body.Font);
        if (style.Header.Color.HasValue)
        {
            range.Style.Border.Left.Style = range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Color.SetColor(style.Header.Color.Value);
            range.Style.Border.Top.Color.SetColor(style.Header.Color.Value);
        }

        
    }
}
