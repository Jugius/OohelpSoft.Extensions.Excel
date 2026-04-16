
namespace OohelpSoft.Excel.DocumentModel;
public readonly struct ExcelCell
{
    public int Row { get; }
    public int Column { get; }


    /// <summary>
    /// Row: 1, Column: 1
    /// </summary>
    public ExcelCell()
    {
        Row = 1;
        Column = 1;
    }
    public ExcelCell(int row, int column)
    {
        this.Row = row;
        this.Column = column;
    }
}
