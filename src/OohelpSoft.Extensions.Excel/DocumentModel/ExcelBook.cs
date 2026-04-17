
namespace OohelpSoft.Excel.DocumentModel;
public class ExcelBook
{
    public ExcelSheets Sheets { get; }
    public ExcelBook() => Sheets = new (this);
    public ExcelBook(IEnumerable<ExcelSheet> sheets) => Sheets = new (this, sheets);    
}
