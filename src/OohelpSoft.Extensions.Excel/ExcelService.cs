using OfficeOpenXml;
using OohelpSoft.Excel.DocumentModel;
using OohelpSoft.Excel.Extensions;
using OohelpSoft.Excel.Interfaces;

namespace OohelpSoft.Excel;
public static class ExcelService
{
    public static IEnumerable<(TValue1, TValue2)> GetValues<TValue1, TValue2>(string filePath, int skipRows = 1)
    {
        using ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath));
        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
        var rows = worksheet.Cells
                    .Select(cell => cell.Start.Row)
                    .Distinct()
                    .OrderBy(x => x);
        var list = rows
            .Skip(skipRows)
            .Select(row => (
                    worksheet.Cells[row, 1].GetValue<TValue1>(),
                    worksheet.Cells[row, 2].GetValue<TValue2>()))
            .ToList();
        return list;
    }
    public static IList<T>Import<T>(FileInfo file, IEnumerable<IPropertySetter<T>> fields) where T : class, new()
    {
        using var excel = new ExcelPackage(file);
        using var workSheet = excel.Workbook.Worksheets[0];

        var collection = workSheet.ConvertSheetToObjects(fields).ToList();
        return collection;
    }
    public static IList<T> Import<T>(byte[] fileBytes, IEnumerable<IPropertySetter<T>> fields) where T : class, new()
    {
        using MemoryStream memoryStream = new MemoryStream(fileBytes, false);
        using ExcelPackage excel = new ExcelPackage(memoryStream);
        using var workSheet = excel.Workbook.Worksheets[0];

        var collection = workSheet.ConvertSheetToObjects(fields).ToList();
        return collection;
    }
    public static async Task<byte[]> Export(ExcelBook book, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(book);

        IProgress<ProgressState>? sheetProgressCounter = null;
        int total = 0;
        int completed = 0;

        if (progress != null)
        {
            total = CalculateTotalRecords(book);
            sheetProgressCounter = new Progress<ProgressState>((e) =>
            {
                progress.Report(new ProgressState(total, ++completed));
            });
        }

        using ExcelPackage package = await Task.Run(() =>
        {

            ExcelPackage package = new ExcelPackage();
            foreach (var sheet in book.Sheets)
            {
                sheet.DrawSheet(package, sheetProgressCounter);
            }

            foreach (var sheet in book.Sheets.Where(a => a.Hidden))
            {
                package.Workbook.Worksheets[sheet.Name].Hidden = eWorkSheetHidden.Hidden;
            }

            return package;

        }).ConfigureAwait(false);

        return await package.GetAsByteArrayAsync().ConfigureAwait(false);
    }
    public static async Task Export(ExcelBook book, Stream stream, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(book);
        ArgumentNullException.ThrowIfNull(stream);

        IProgress<ProgressState>? sheetProgressCounter = null;
        int total = 0;
        int completed = 0;

        if (progress != null)
        {
            total = CalculateTotalRecords(book);
            sheetProgressCounter = new Progress<ProgressState>((e) =>
            {
                progress.Report(new ProgressState(total, ++completed));
            });
        }

        using ExcelPackage package = await Task.Run(() =>
        {

            ExcelPackage package = new ExcelPackage();
            foreach (var sheet in book.Sheets)
            {
                sheet.DrawSheet(package, sheetProgressCounter);
            }

            foreach (var sheet in book.Sheets.Where(a => a.Hidden))
            {
                package.Workbook.Worksheets[sheet.Name].Hidden = eWorkSheetHidden.Hidden;
            }

            return package;

        }).ConfigureAwait(false);

        await package.SaveAsAsync(stream).ConfigureAwait(false);
    }
    public static ExcelPackage CreatePackage(ExcelBook book, IProgress<ProgressState>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(book);

        IProgress<ProgressState>? sheetProgressCounter = null;
        int total = 0;
        int completed = 0;

        if (progress != null)
        {
            if (book.Sheets.Count > 1)
            {
                total = CalculateTotalRecords(book);
                sheetProgressCounter = new Progress<ProgressState>((e) =>
                {
                    progress.Report(new ProgressState(total, ++completed));
                });
            }
            else 
            {
                sheetProgressCounter = progress;
            }            
        }

        ExcelPackage package = new ExcelPackage();
        foreach (var sheet in book.Sheets)
        {
            sheet.DrawSheet(package, sheetProgressCounter);
        }

        foreach (var sheet in book.Sheets.Where(a => a.Hidden))
        {
            package.Workbook.Worksheets[sheet.Name].Hidden = eWorkSheetHidden.Hidden;
        }

        return package;
    }
    
    private static int CalculateTotalRecords(ExcelBook book)
    {
        int total = 0;
        foreach (var sheet in book.Sheets)
        {
            total += sheet.Table!.RowsCout;
        }
        return total;
    }
}
