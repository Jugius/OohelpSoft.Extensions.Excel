
namespace OohelpSoft.Excel.DocumentModel.Styles;
public class ExcelTableStyle
{
    public ExcelTableHeaderStyle Header { get; } = new ExcelTableHeaderStyle();
    public ExcelTableBodyStyle Body { get; } = new ExcelTableBodyStyle();
    
}
