using OfficeOpenXml;

namespace OohelpSoft.Excel.Utils;
public static class Hyperlink
{
    private const int MaxExcelUriLength = 255;        
    public static void DrawHyperlink(this ExcelRange cell, Uri uri, string cellText, bool asFormula = false)
    {
        ArgumentNullException.ThrowIfNull(cell);

        if (uri == null || !uri.IsAbsoluteUri) return;

        if (asFormula)
        {
            DrawFormula(cell, uri, cellText);
            return;
        }

        cell.Hyperlink = uri;
        cell.Value = cellText;

        cell.Style.Font.UnderLine = true;
        cell.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Hyperlink);
    }
    public static void DrawFormula(ExcelRange cell, Uri uri, string cellText)
    {
        ArgumentNullException.ThrowIfNull(cell);

        if (uri == null) return;
        if (TryGetNormalizedExcelUri(uri, out string normalized))
        {
            cell.Formula = "HYPERLINK(\"" + normalized + "\",\"" + cellText + "\")";
            cell.Style.Font.UnderLine = true;
            cell.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Hyperlink);
        }
        else
        {
            cell.Value = uri.ToString();
        }
    }
    public static Uri? GetUri(this ExcelRange cell)
    {
        if (cell == null)
            return null;

        if (cell.Hyperlink != null)
            return cell.Hyperlink;


        if (cell.Value != null)
        {
            string url = cell.GetValue<string>().Trim();
            if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out Uri? u))
                return u;
        }

        if (cell.Formula != null)
        {
            string formula = cell.Formula;
            if (formula.Contains("HYPERLINK", StringComparison.OrdinalIgnoreCase))
            {
                int Start, End;
                Start = formula.IndexOf('"', 0) + 1;
                End = formula.IndexOf('"', Start);
                formula = formula[Start..End];
                if (Uri.TryCreate(Uri.UnescapeDataString(formula), UriKind.Absolute, out Uri? u))
                    return u;
            }
        }
        return null;
    }
    private static bool TryGetNormalizedExcelUri(Uri uri, out string normalized)
    {
        normalized = string.Empty;

        if (uri == null || !uri.IsAbsoluteUri)
            return false;

        var original = uri.OriginalString;
                
        if (original.Contains('"', StringComparison.OrdinalIgnoreCase) || original.Contains('\'', StringComparison.OrdinalIgnoreCase))
            return false;
        
        var candidate = original.Replace(" ", "%20", StringComparison.OrdinalIgnoreCase);
                
        if (candidate.Length > MaxExcelUriLength)
            return false;

        if (!Uri.TryCreate(candidate, UriKind.Absolute, out _))
            return false;

        normalized = candidate;
        return true;
    } 
}
