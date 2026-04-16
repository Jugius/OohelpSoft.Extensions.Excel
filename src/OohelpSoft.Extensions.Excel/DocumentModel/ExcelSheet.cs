using OfficeOpenXml;

namespace OohelpSoft.Excel.DocumentModel;

public class ExcelSheet
{
    private readonly ExcelBook parent;

    public ExcelBook Parent => parent;
    public string Name { get; set; } = "Лист1";
    public bool Hidden { get; set; }
    public bool ShowGridLines { get; set; }
    public bool CheckHyperlinksDrawing { get; set; } = true;
    public ExcelTable? Table { get; set; }

    public virtual ExcelWorksheet DrawSheet(ExcelPackage excel, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(excel);

        ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(this.Name);
        worksheet.View.ShowGridLines = this.ShowGridLines;

        if(this.CheckHyperlinksDrawing)
            this.CheckHyperlinks();

        this.Table?.DrawTable(worksheet, progress);
        return worksheet;
    }
    public ExcelSheet(ExcelBook parent)
    {
        this.parent = parent;
    }
    public ExcelTable<T> CreateTable<T>(IList<T> items, string tableName) where T : class
    {
        ExcelTable<T> table = new (items, this)
        {
            Name = tableName
        };

        this.Table = table;
        return table;
    }
    private void CheckHyperlinks()
    {
        var hyperlinkFields = this.Table!.Fields.OfType<Interfaces.IExcelHyperlink>().ToList();
        int potentialLinksCount = hyperlinkFields.Count * this.Table.RowsCout;

        if (potentialLinksCount <= Constants.Sheet.MaxHyperlinks) return;

        foreach (var c in hyperlinkFields)
        {
            c.DrawAsFormula = true;
        }
    }
}
