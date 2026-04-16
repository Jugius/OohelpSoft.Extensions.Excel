
namespace OohelpSoft.Excel.DocumentModel;
public class ExcelBook
{
    private readonly List<ExcelSheet> sheets;

    public IList<ExcelSheet> Sheets => sheets;
    public ExcelBook()
    {
        this.sheets = [];
    }
    public ExcelBook(IEnumerable<ExcelSheet> sheets)
    {
        this.sheets = [.. sheets];
    }
    public ExcelSheet CreateSheet(string sheetName)
    {
        ExcelSheet sheet = new (this)
        {
            Name = sheetName
        };

        this.sheets.Add(sheet);
        return sheet;
    }    
}
