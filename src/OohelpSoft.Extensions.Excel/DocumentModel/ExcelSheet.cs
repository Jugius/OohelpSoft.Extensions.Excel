using OfficeOpenXml;

namespace OohelpSoft.Excel.DocumentModel;

public class ExcelSheet
{
    public ExcelBook Parent { get; }
    public string Name { get; set; } = "Лист1";
    public bool Hidden { get; set; }
    public bool ShowGridLines { get; set; }
    public bool CheckHyperlinksDrawing { get; set; } = true;
    public ExcelTables Tables { get; }   

    public virtual ExcelWorksheet DrawSheet(ExcelPackage excel, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(excel);

        ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(this.Name);
        worksheet.View.ShowGridLines = this.ShowGridLines;

        if(this.CheckHyperlinksDrawing)
            this.CheckHyperlinks();

        foreach (var table in this.Tables)
        {
            table.DrawTable(worksheet, progress);
        }
        
        return worksheet;
    }
    public ExcelSheet(ExcelBook parent)
    {
        this.Parent = parent;
        this.Tables = new(this);
    }
    private void CheckHyperlinks()
    {
        var hyperlinkFields = this.Tables.SelectMany(t => t.Fields.OfType<Interfaces.IExcelHyperlink>()).ToList();
        int potentialLinksCount = hyperlinkFields.Count * this.Tables.Sum(t => t.RowsCout);

        if (potentialLinksCount <= Constants.Sheet.MaxHyperlinks) return;

        foreach (var c in hyperlinkFields)
        {
            c.DrawAsFormula = true;
        }
    }
}
