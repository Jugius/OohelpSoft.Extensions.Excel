using System.Collections;

namespace OohelpSoft.Excel.DocumentModel;

public sealed class ExcelFields : IEnumerable<ExcelField>
{
    private readonly List<ExcelField> _fields = [];
    public int Count => _fields.Count;

    public void Add(ExcelField field) => _fields.Add(field);
    public void AddRange(IEnumerable<ExcelField> fields) => _fields.AddRange(fields);

    public IEnumerator<ExcelField> GetEnumerator() => _fields.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
