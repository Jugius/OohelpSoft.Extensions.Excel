using System.Drawing;

namespace OohelpSoft.Excel.DocumentModel.Styles;
public class ExcelTableHeaderStyle
{
    public ExcelFont Font { get; } = new ExcelFont { Color = System.Drawing.Color.FromArgb(216, 231, 249) };
    public Color? Color { get; set; } = System.Drawing.Color.FromArgb(0, 83, 127);
    public bool WrapText { get; set; } = true;
    public double? Height { get; set; }
}
