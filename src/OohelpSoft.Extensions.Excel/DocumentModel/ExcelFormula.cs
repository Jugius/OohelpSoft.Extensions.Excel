using OfficeOpenXml;

namespace OohelpSoft.Excel.DocumentModel;
public abstract class ExcelFormula : ExcelField
{
    private string? _formula;
    public virtual string Formula => _formula ??= InitializeFormula();
    protected ExcelFormula()
    {
        this.NumberFormat = Constants.NumberFormat.FormattedInteger;
    }
    public override void DrawCell<TSource>(TSource source, ExcelRange cell)
    {
        ArgumentNullException.ThrowIfNull(cell);
        cell.Formula = Formula;
    }
    protected abstract string InitializeFormula();
}
