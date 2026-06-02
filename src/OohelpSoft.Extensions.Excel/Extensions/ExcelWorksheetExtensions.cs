using OfficeOpenXml;
using OohelpSoft.Excel.Interfaces;

namespace OohelpSoft.Excel.Extensions;
internal static class ExcelWorksheetExtensions
{    
    public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet, IEnumerable<IPropertySetter<T>> columns, int startRow = 2) where T : class, new()
    {
        var rows = worksheet.Cells
            .Select(cell => cell.Start.Row)
            .Distinct()
            .OrderBy(x => x);

        
        var collection = rows.Skip(startRow - 1).Select(row =>
        {
            T tnew = new T();
            int columnPos = 1;

            foreach (var column in columns)
            {
                var cell = worksheet.Cells[row, columnPos];
                column.SetPropertyValue(cell, tnew);
                columnPos++;
            }
            return tnew;
        });

        //Send it back
        return collection;
    }    
}
