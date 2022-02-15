using Excel_To_HTML;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

const string ExcelToPdfSerialNumber = "50013243563";
const string DocumentNetSerialNumber = "50024532950";
const string FileName = "SampleDoc";
const double DefaultInnerXPadding = 0.0;
const double DefaultInnerYPadding = 1.0;
const double DefaultFrameXPadding = 10.0;
const double DefaultFrameYPadding = 8.0;
const int DefaultFontSize = 12;
const int DefaultRowCount = 12;

DocumentCore.Serial = DocumentNetSerialNumber;

try
{
    var pdf = await ConvertToPdfAsync();

    await ConvertToHtmlAsync(pdf);

    Console.WriteLine("Done.");
}
catch (Exception ex)
{
    Console.WriteLine(ex);
}

async Task<byte[]> ConvertToPdfAsync()
{
    var excel = await File.ReadAllBytesAsync($".\\{FileName}.xlsx");

    var excelToPdf = new SautinSoft.ExcelToPdf
    {
        Serial = ExcelToPdfSerialNumber,
        OutputFormat = SautinSoft.ExcelToPdf.eOutputFormat.Pdf,
    };

    excelToPdf.PageStyle.PageOrientation.Landscape();
    excelToPdf.PageStyle.PageScale.FitByWidth();
    excelToPdf.PageStyle.PageSize.A4();

    var pdf = excelToPdf.ConvertBytes(excel);

    await File.WriteAllBytesAsync($".\\{FileName}.pdf", pdf);

    return pdf;
}

async Task ConvertToHtmlAsync(byte[] pdf)
{
    var document = LoadDocument(pdf);
    var certificate = new PrintformCertificate
    {
        OrganizationName = "ООО \"Ромашка\"",
        FirstName = "Семен",
        LastName = "Горбунков",
        Surname = "Семенович",
        SerialNumber = "123",
        ValidFrom = DateTime.Now,
        ValidTo = DateTime.Now
    };

    var sign = new Sign(certificate, DateTime.Now, DateTime.Now);

    AddSignature2(document, new[] { sign });

    using (var ms = new MemoryStream())
    {
        document.Save(ms, SaveOptions.HtmlFixedDefault);

        var html = ms.ToArray();

        await File.WriteAllBytesAsync($".\\{FileName}.html", html);
    }
}

DocumentCore LoadDocument(byte[] pdf)
{
    using (var ms = new MemoryStream(pdf))
    {
        return DocumentCore.Load(ms, LoadOptions.PdfDefault);
    }
}

void AddSignature2(DocumentCore document, IReadOnlyList<Sign> signatures)
{
    if (signatures.Count == 0)
    {
        // подписей нет, дорисовывать нечего
        return;
    }

    // ищем текущую последнюю секцию
    var lastSection = document.Sections.LastOrDefault();
    if (lastSection == null)
    {
        // если документ пустой, ничего не делаем
        return;
    }

    // добавляем новую секцию с такими же настройками страницы, как и у предыдущей секции
    var section = new Section(document)
    {
        PageSetup = lastSection.PageSetup.Clone()
    };

    // теперь нам нужно убедиться, что в новой секции единственная текстовая колонка,
    section.PageSetup.TextColumns = new TextColumnCollection(1);

    // и она не выводится на титульной странице
    section.PageSetup.TitlePage = false;

    for (var i = 0; i < signatures.Count; i++)
    {
        // если это первая подпись, то гарантируем, что новая секция будет выведена на новой странице;
        // для последующих подписей обесспечиваем пустую строку между текущей подписбью и предыдущей
        var newLine = new Paragraph(document, new SpecialCharacter(document, i == 0 ? SpecialCharacterType.PageBreak : SpecialCharacterType.LineBreak));
        section.Blocks.Add(newLine);

        // рисуем таблицу с подписями
        var table = GetSignatureTable(document, signatures[i]);
        section.Blocks.Add(table);
    }

    // добавляем новую секцию в документ
    document.Sections.Add(section);
}

Table GetSignatureTable(DocumentCore document, Sign signature)
{
    // рисуем таблицу-рамку
    var frameTable = GetFrameTable(document);

    // риусем внутреннюю таблицу
    var innerTable = GetInnerTable(document, signature);

    // вставляем внутреннюю таблицу в рамку
    frameTable.Rows[0].Cells[0].Blocks.Add(innerTable);

    return frameTable;
}

Table GetFrameTable(DocumentCore document)
{
    // рисуем таблицу-рамку из одной ячейки
    var table = new Table(document);
    var row = new TableRow(document);
    var cell = new TableCell(document);

    // у ячейки будут внешние границы, и она будет занимать всю ширину таблицы
    cell.CellFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Blue, 1);
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
    cell.CellFormat.Padding = new Padding(DefaultFrameXPadding, DefaultFrameYPadding, DefaultFrameXPadding, DefaultFrameYPadding, LengthUnit.Millimeter);

    row.Cells.Add(cell);
    table.Rows.Add(row);

    // растягиваем таблицу на всю ширину родительской секции
    table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
    table.TableFormat.Alignment = HorizontalAlignment.Center;

    return table;
}

Table GetInnerTable(DocumentCore document, Sign signature)
{
    // определяем число строк в таблице
    var rowCount = DefaultRowCount;

    if (!string.IsNullOrEmpty(signature.SigningTime))
    {
        // строка для вывода времени подписания
        rowCount++;
    }

    if (!string.IsNullOrEmpty(signature.SignatureTimeStampTime))
    {
        // строка для вывода штампа времени, заверающего время подписания
        rowCount++;
    }

    // рисуем таблицу
    var table = new Table(document);

    for (var i = 0; i < rowCount; i++)
    {
        var row = i switch
        {
            0 => GetInnerTableRowOrDefault(document, "Подписано электронной подписью", signature.Employee),
            1 => GetInnerTableRowOrDefault(document, "Серийный номер сертификата", signature.SerialNumber),
            2 => GetInnerTableRowOrDefault(document, "Период действия сертификата", signature.ValidityPeriod),
            3 => GetInnerTableRowOrDefault(document, "Штамп времени", signature.SigningTime, MultipleBorderTypes.Top),
            4 => GetInnerTableRowOrDefault(document, "Подпись заверена", signature.SignatureTimeStampTime),
            _ => throw new NotImplementedException()
        };

        if (row != null)
        {
            table.Rows.Add(row);
        }
    }

    // растягиваем таблицу до 90% родительской ячейки
    table.TableFormat.PreferredWidth = new TableWidth(90, TableWidthUnit.Percentage);
    table.TableFormat.Alignment = HorizontalAlignment.Center;

    return table;
}

TableRow? GetInnerTableRowOrDefault(DocumentCore document, string left, string right, MultipleBorderTypes borderTypes = MultipleBorderTypes.None)
{
    if (string.IsNullOrEmpty(right))
    {
        return null;
    }

    var row = new TableRow(document);
    row.Cells.Add(GetInnerTableCell(document, borderTypes, left));
    row.Cells.Add(GetInnerTableCell(document, borderTypes, right));
    return row;
}

TableCell GetInnerTableCell(DocumentCore document, MultipleBorderTypes borderTypes, string text)
{
    var block = new Paragraph(document, new Run(document, text, new CharacterFormat
    {
        FontColor = Color.Blue,
        Size = DefaultFontSize
    }));

    var cell = new TableCell(document, block);

    cell.CellFormat.PreferredWidth = new TableWidth(50, TableWidthUnit.Percentage);
    cell.CellFormat.Borders.Add(borderTypes, BorderStyle.Single, Color.Blue, 1.0);
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
    cell.CellFormat.Padding = new Padding(DefaultInnerXPadding, DefaultInnerYPadding, DefaultInnerXPadding, DefaultInnerYPadding, LengthUnit.Millimeter);

    return cell;
}
