using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace OohelpSoft.Excel.DocumentModel;

public abstract class ExcelField
{
    public virtual string? NumberFormat { get; set; }
    public virtual string ColumnHeader { get; set; } = null!;
    public virtual int ColumnWidth { get; set; }
    public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.Left;
    public abstract void DrawCell<T>(T source, ExcelRange cell);
}
public abstract class ExcelField<TSource> : ExcelField where TSource : class
{
    public abstract void DrawCell(TSource source, ExcelRange cell);
    public override void DrawCell<T>(T source, ExcelRange cell)
    {
        if (source is TSource tsou)
            DrawCell(tsou, cell);
        else
            throw new ArgumentException($"Invalid source type. Expected {typeof(TSource)}, but got {typeof(T)}");
    }
    protected static int? NullableOrNull(int? val) => val.HasValue && val.Value != 0 ? val.Value : null;
#pragma warning disable S1244 // Floating point numbers should not be tested for equality
    protected static double? NullableOrNull(double? val) => val.HasValue && val.Value != 0 ? val.Value : null;
    protected static float? NullableOrNull(float? val) => val.HasValue && val.Value != 0 ? val.Value : null;
#pragma warning restore S1244 // Floating point numbers should not be tested for equality
}