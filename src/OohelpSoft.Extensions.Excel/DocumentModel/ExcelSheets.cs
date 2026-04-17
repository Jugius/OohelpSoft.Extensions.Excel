using System.Collections;

namespace OohelpSoft.Excel.DocumentModel;

public class ExcelSheets : IEnumerable<ExcelSheet>
{
    private readonly List<ExcelSheet> _sheets;
    private readonly ExcelBook _book;
    public int Count => _sheets.Count;
    internal ExcelSheets(ExcelBook book)
    { 
        _book = book;
        _sheets = [];
    }
    internal ExcelSheets(ExcelBook book, IEnumerable<ExcelSheet> sheets)
    {
        _book = book;
        _sheets = [.. sheets];
    }
    public ExcelSheet Add(string sheetName)
    {
        ExcelSheet sheet = new(_book)
        {
            Name = sheetName
        };

        _sheets.Add(sheet);
        return sheet;
    }
    public TSheet Add<TSheet>(string sheetName, Func<ExcelBook, TSheet> factory)
        where TSheet : ExcelSheet
    {
        ArgumentNullException.ThrowIfNull(factory);

        var sheet = factory(_book);
        sheet.Name = sheetName;

        _sheets.Add(sheet);
        return sheet;
    }

    public IEnumerator<ExcelSheet> GetEnumerator() => _sheets.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _sheets.GetEnumerator();
}
