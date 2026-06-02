using OfficeOpenXml;
using OfficeOpenXml.Table;
using OohelpSoft.Excel.Extensions;

namespace OohelpSoft.Excel.DocumentModel;

public class ExcelSmartTable<TExcelItem> : ExcelTable<TExcelItem> where TExcelItem : class
{
    public string Name { get; }
    public TableStyles TableStyle { get; set; } = TableStyles.Light9;
    public ExcelSmartTable(ExcelSheet parent, IList<TExcelItem> items, string name) : base(parent, items)
    {
        Name = name;
    }
    public override void DrawTable(ExcelWorksheet worksheet, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        using var table = DrawTableStructure(worksheet);        
        OfficeOpenXml.Table.ExcelTable tab = worksheet.Tables.Add(table, this.Name);
        tab.TableStyle = OfficeOpenXml.Table.TableStyles.Light9;

        table.ApplyTableStyles(this.Style);

        FillTableWithData(worksheet, progress);
    }
}
