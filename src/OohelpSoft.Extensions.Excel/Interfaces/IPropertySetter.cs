using OfficeOpenXml;

namespace OohelpSoft.Excel.Interfaces;
public interface IPropertySetter<TDestination> where TDestination : notnull, new()
{
    void SetPropertyValue(ExcelRange cell, TDestination destination);
}
