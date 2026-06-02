using OfficeOpenXml;
using OohelpSoft.Excel.DocumentModel.Styles;
using OohelpSoft.Excel.Extensions;

namespace OohelpSoft.Excel.DocumentModel;
public abstract class ExcelTable
{   
    private readonly ExcelSheet parentSheet;
    public ExcelSheet Parent => parentSheet;
    public ExcelTableStyle Style { get; } = new ExcelTableStyle();
    public ExcelCell TopLeftCell { get; set; } = new ExcelCell();
    public ExcelFields Fields { get; } = new();
    public abstract int RowsCout { get; }
    protected ExcelTable(ExcelSheet parent) => this.parentSheet = parent;
    public abstract void DrawTable(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null);
    protected virtual ExcelRange DrawTableStructure(ExcelWorksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        int firstRow = TopLeftCell.Row;
        int firstCol = TopLeftCell.Column;
        int lastCol = firstCol + Fields.Count - 1;
        int lastRow = firstRow + (RowsCout > 0 ? RowsCout : 1);

        int colIndex = firstCol;
        foreach (var field in Fields)
        {
            worksheet.Cells[firstRow, colIndex].Value = field.ColumnHeader;
            var column = worksheet.Column(colIndex);
            column.Style.HorizontalAlignment = field.HorizontalAlignment;
            column.Width = field.ColumnWidth;
            if (field.NumberFormat != null)
                column.Style.Numberformat.Format = field.NumberFormat;

            colIndex++;
        }

        return worksheet.Cells[firstRow, firstCol, lastRow, lastCol];        
    }
}
public class ExcelTable<TExcelItem> : ExcelTable where TExcelItem : class
{
    public ExcelTable(ExcelSheet parent, IList<TExcelItem> items) : base(parent)
    {
        this.Items = items;
    }
    public IList<TExcelItem> Items { get; }
    public override int RowsCout => this.Items.Count;

    public override void DrawTable(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        using var table = DrawTableStructure(worksheet);
        table.ApplyTableStyles(this.Style);

        FillTableWithData(worksheet, progress);        
    }

    

    protected virtual void FillTableWithData(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

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
