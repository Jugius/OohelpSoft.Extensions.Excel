
namespace OohelpSoft.Excel.DocumentModel;
public readonly struct ExcelCell : IEquatable<ExcelCell>
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

    public override bool Equals(object? obj) => obj is ExcelCell cell && this.Equals(cell);
    public bool Equals(ExcelCell other) => this.Row == other.Row && this.Column == other.Column;
    public override int GetHashCode() => HashCode.Combine(Row, Column);

    public static bool operator ==(ExcelCell left, ExcelCell right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(ExcelCell left, ExcelCell right)
    {
        return !(left == right);
    }
}
