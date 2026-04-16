using OfficeOpenXml;

namespace OohelpSoft.Excel.Interfaces;
public interface IPropertySetter<in TDestination> where TDestination : notnull, new()
{
    void SetPropertyValue(ExcelRange cell, TDestination destination);
}
