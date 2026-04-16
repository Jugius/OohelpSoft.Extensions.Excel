using OfficeOpenXml;
using OohelpSoft.Excel.DocumentModel;
using OohelpSoft.Excel.Interfaces;

namespace OohelpSoft.Excel.Extensions;
internal static class ExcelWorksheetExtensions
{    
    public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet, IEnumerable<IPropertySetter<T>> columns) where T : class, new()
    {
        var rows = worksheet.Cells
            .Select(cell => cell.Start.Row)
            .Distinct()
            .OrderBy(x => x);

        
        var collection = rows.Skip(1).Select(row =>
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

    public static ExcelRange AddTable(this ExcelWorksheet worksheet, ExcelTable table)
    {
        int firstRow = table.TopLeftCell.Row;
        int firstCol = table.TopLeftCell.Column;
        int lastCol = firstCol + table.Fields.Count - 1;
        int lastRow = firstRow + (table.RowsCout > 0 ? table.RowsCout : 1);

        int colIndex = firstCol;
        foreach (var field in table.Fields)
        {
            worksheet.Cells[firstRow, colIndex].Value = field.ColumnHeader;
            var column = worksheet.Column(colIndex);
            column.Style.HorizontalAlignment = field.HorizontalAlignment;
            column.Width = field.ColumnWidth;
            if (field.NumberFormat != null)
                column.Style.Numberformat.Format = field.NumberFormat;

            colIndex++;
        }
        
        ExcelRange tableRange = worksheet.Cells[firstRow, firstCol, lastRow, lastCol];        
        OfficeOpenXml.Table.ExcelTable tab = worksheet.Tables.Add(tableRange, table.Name);
        tab.TableStyle = OfficeOpenXml.Table.TableStyles.Light9;
        return tableRange;
    }
    
}
