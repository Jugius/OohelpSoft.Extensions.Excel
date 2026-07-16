using System.IO.Compression;
using System.Xml.Linq;

namespace OohelpSoft.Excel.DocumentModel;

public class XlsxProperties
{
    // Пространства имен OpenXML
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Ap = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    public string? Title { get; set; }
    public string? Subject {  get; set; }
    public string? Author {  get; set; }
    public string? Comments { get; set; } = string.Empty;
    public string? Keywords { get; set; } = string.Empty;
    public string? Category { get; set; }
    public string? Company { get; set; } = string.Empty;
    public string? Application { get; set; } = string.Empty;
    public string? AppVersion { get; set; }

    public static void UpdateXlsx(
        Stream stream, XlsxProperties properties)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(properties);

        if (!stream.CanRead || !stream.CanWrite || !stream.CanSeek)
        {
            throw new ArgumentException("Поток должен поддерживать операции Read, Write и Seek (позиционирование).", nameof(stream));
        }

        // Возвращаем указатель в начало потока, так как EPPlus записал данные и ушел в конец
        stream.Position = 0;

        // Открываем ZipArchive прямо поверх стрима.
        // leaveOpen: true критически важен, чтобы ZipArchive не закрыл наш stream при закрытии!
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            // 1. Обработка core.xml
            var coreEntry = archive.GetEntry("docProps/core.xml");
            if (coreEntry != null)
            {
                XDocument doc;
                using (var entryStream = coreEntry.Open())
                {
                    doc = XDocument.Load(entryStream);
                }

                var root = doc.Root;
                if (root != null)
                {
                    bool isModified = false;

                    isModified |= TryUpdateElement(root, Dc + "title", properties.Title);
                    isModified |= TryUpdateElement(root, Dc + "subject", properties.Subject);
                    isModified |= TryUpdateElement(root, Dc + "creator", properties.Author);
                    isModified |= TryUpdateElement(root, Dc + "description", properties.Comments);
                    isModified |= TryUpdateElement(root, Cp + "keywords", properties.Keywords);
                    isModified |= TryUpdateElement(root, Cp + "category", properties.Category);

                    if (isModified)
                    {
                        using var entryStream = coreEntry.Open();
                        entryStream.SetLength(0); // Стираем старый XML
                        doc.Save(entryStream);
                    }
                }
            }

            // 2. Обработка app.xml
            var appEntry = archive.GetEntry("docProps/app.xml");
            if (appEntry != null)
            {
                XDocument doc;
                using (var entryStream = appEntry.Open())
                {
                    doc = XDocument.Load(entryStream);
                }

                var root = doc.Root;
                if (root != null)
                {
                    bool isModified = false;

                    isModified |= TryUpdateElement(root, Ap + "Company", properties.Company);
                    isModified |= TryUpdateElement(root, Ap + "Application", properties.Application);
                    isModified |= TryUpdateElement(root, Ap + "AppVersion", properties.AppVersion);

                    if (isModified)
                    {
                        using var entryStream = appEntry.Open();
                        entryStream.SetLength(0); // Стираем старый XML
                        doc.Save(entryStream);
                    }
                }
            }
        } // Здесь ZipArchive сохраняет изменения в поток (так как Update), но НЕ закрывает сам stream.

        // Сбрасываем позицию в 0, чтобы тот, кто вызвал Export, мог сразу читать стрим с начала
        stream.Position = 0;
    }

    private static bool TryUpdateElement(XElement parent, XName name, string? newValue)
    {
        if (newValue == null)
            return false;

        var element = parent.Element(name);
        if (element != null)
        {
            if (element.Value != newValue)
            {
                element.Value = newValue;
                return true;
            }
        }
        else
        {
            parent.Add(new XElement(name, newValue));
            return true;
        }

        return false;
    } 
}
