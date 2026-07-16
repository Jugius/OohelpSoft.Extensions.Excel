
using OfficeOpenXml;

namespace OohelpSoft.Excel.DocumentModel;
public class ExcelBook
{
    public XlsxProperties Properties { get; } = new XlsxProperties();
    public ExcelSheets Sheets { get; }
    public ExcelBook() => Sheets = new (this);
    public ExcelBook(IEnumerable<ExcelSheet> sheets) => Sheets = new (this, sheets);

    public async Task<MemoryStream> ToStreamAsync(IProgress<ProgressState>? progress = null)
    {
        var memoryStream = new MemoryStream();
        await SaveTo(memoryStream, progress).ConfigureAwait(false);        
        return memoryStream;
    }

    public async Task SaveTo(Stream stream, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using (ExcelPackage package = await ExcelService.CreatePackageAsync(this, progress).ConfigureAwait(false))
        {
            await package.SaveAsAsync(stream).ConfigureAwait(false);
        }

        XlsxProperties.UpdateXlsx(stream, this.Properties);
    }
}
