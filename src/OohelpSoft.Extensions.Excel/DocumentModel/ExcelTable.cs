using OfficeOpenXml;
using OohelpSoft.Excel.DocumentModel.Styles;
using OohelpSoft.Excel.Extensions;

namespace OohelpSoft.Excel.DocumentModel;
public abstract class ExcelTable
{   
    private readonly ExcelSheet parentSheet;
    public ExcelSheet Parent => parentSheet;
    public ExcelTableStyle Style { get; } = new ExcelTableStyle();
    public string Name {  get; set; } = null!;
    public ExcelCell TopLeftCell { get; set; } = new ExcelCell();
    public ExcelFields Fields { get; } = new();
    public abstract int RowsCout { get; }
    protected ExcelTable(ExcelSheet parent) => this.parentSheet = parent;
    public abstract void DrawTable(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null);    
}
public class ExcelTable<TExcelItem>(IList<TExcelItem> items, ExcelSheet parent) : ExcelTable(parent) where TExcelItem : class
{
    public IList<TExcelItem> Items { get; } = items;
    public override int RowsCout => this.Items.Count;

    public override void DrawTable(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        using var table = worksheet.AddTable(this);
        table.ApplyTableStyles(this.Style);

        int row = this.TopLeftCell.Row + 1;
        int total = this.Items.Count;
        int completed = 0;
        foreach (TExcelItem item in this.Items)
        {
            int column = this.TopLeftCell.Column;
            foreach (var field in this.Fields)
            {
                using var cell = worksheet.Cells[row, column];
                field.DrawCell(item, cell);
                column++;
            }
            row++;
            progress?.Report(new ProgressState(total, ++completed));
        }
    }
}
