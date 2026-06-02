
using System.Collections;

namespace OohelpSoft.Excel.DocumentModel;

public sealed class ExcelTables : IEnumerable<ExcelTable>
{
    private readonly List<ExcelTable> _tables;
    private readonly ExcelSheet _sheet;

    internal ExcelTables(ExcelSheet sheet)
    {
        _sheet = sheet;
        _tables = [];
    }

    public ExcelTable Add<TTable>(Func<ExcelSheet, TTable> factory)
        where TTable : ExcelTable
    {
        ArgumentNullException.ThrowIfNull(factory);

        var table = factory(_sheet);
        _tables.Add(table);

        return table;
    }
    public ExcelTable<T> Add<T>(IList<T> items) where T : class
    {
        var table = new ExcelTable<T>(this._sheet, items);
        _tables.Add(table);

        return table;
    }
    public ExcelSmartTable<TExcelItem> Add<TExcelItem>(IList<TExcelItem> items, string tableName) where TExcelItem : class
    {
        var table = new ExcelSmartTable<TExcelItem>(this._sheet, items, tableName);
        _tables.Add(table);

        return table;
    }
    public IEnumerator<ExcelTable> GetEnumerator() => _tables.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _tables.GetEnumerator();
}
