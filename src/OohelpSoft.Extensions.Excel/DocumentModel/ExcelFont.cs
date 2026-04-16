using System.Drawing;


namespace OohelpSoft.Excel.DocumentModel;
public class ExcelFont
{
    public string Name { get; set; } = "Calibri";
    public float Size { get; set; } = 10F;
    public bool Bold {  get; set; } = false;
    public bool Italic { get; set; } = false;
    public Color? Color { get; set; } = System.Drawing.Color.FromArgb(30, 55, 55);
}
