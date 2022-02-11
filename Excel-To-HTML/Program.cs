using Excel_To_HTML;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

const string ExcelToPdfSerialNumber = "50013243563";
const string DocumentNetSerialNumber = "50024532950";
const string FileName = "SampleDoc";
const int DefaultMargin = 10;
const double DefaultLeftPadding = 2.0;
const double DefaultPadding = 0.5;
const int DefaultFontSize = 12;

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

    excelToPdf.PageStyle.PageMarginLeft.mm(DefaultMargin);
    excelToPdf.PageStyle.PageMarginTop.mm(DefaultMargin);
    excelToPdf.PageStyle.PageMarginRight.mm(DefaultMargin);
    excelToPdf.PageStyle.PageMarginBottom.mm(DefaultMargin);
    excelToPdf.PageStyle.PageOrientation.Landscape();
    excelToPdf.PageStyle.PageScale.Auto();
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
    // определяем число строк в таблице
    var rowCount = 3;

    if (!string.IsNullOrEmpty(signature.SigningTime))
    {
        rowCount++;
    }
    
    if (!string.IsNullOrEmpty(signature.SignatureTimeStampTime))
    {
        rowCount++;
    }

    // риусем таблицу с подписью
    var table = new Table(document);
    
    for (var i = 0; i < rowCount; i++)
    {
        switch (i)
        {
            case 0:
                {
                    AddSigner(document, table, signature);
                    break;
                }

            case 1:
                {
                    AddCertificateSerial(document, table, signature);
                    break;
                }

            case 2:
                {
                    AddCertificateValidityPeriod(document, table, signature);
                    break;
                }

            case 3:
                {
                    AddSigningTimestamp(document, table, signature);
                    break;
                }

            case 4:
                {
                    if (!string.IsNullOrEmpty(signature.SignatureTimeStampTime))
                    {
                        AddSignatureTimeStampTime(document, table, signature);
                    }
                    break;
                }
        }
    }

    // растягиваем таблицу на всю ширину секции
    table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);

    return table;
}

void AddSigner(DocumentCore document, Table table, Sign signature)
{
    var row = new TableRow(document);
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Left | MultipleBorderTypes.Top, "Подписано электронной подписью"));
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Right | MultipleBorderTypes.Top, signature.Employee));
    table.Rows.Add(row);
}

void AddCertificateSerial(DocumentCore document, Table table, Sign signature)
{
    var row = new TableRow(document);
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Left, "Серийный номер сертификата"));
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Right, signature.SerialNumber));
    table.Rows.Add(row);
}

void AddCertificateValidityPeriod(DocumentCore document, Table table, Sign signature)
{
    var leftBorder = MultipleBorderTypes.Left;
    var rightBorder = MultipleBorderTypes.Right;

    if (string.IsNullOrEmpty(signature.SigningTime))
    {
        leftBorder |= MultipleBorderTypes.Bottom;
        rightBorder |= MultipleBorderTypes.Bottom;
    }

    var row = new TableRow(document);
    row.Cells.Add(GetSignerTableCell(document, leftBorder, "Период действия сертификата"));
    row.Cells.Add(GetSignerTableCell(document, rightBorder, signature.ValidityPeriod));
    table.Rows.Add(row);
}

void AddSigningTimestamp(DocumentCore document, Table table, Sign signature)
{
    var leftBorder = MultipleBorderTypes.Left | MultipleBorderTypes.Top;
    var rightBorder = MultipleBorderTypes.Right | MultipleBorderTypes.Top;

    if (string.IsNullOrEmpty(signature.SignatureTimeStampTime))
    {
        leftBorder |= MultipleBorderTypes.Bottom;
        rightBorder |= MultipleBorderTypes.Bottom;
    }

    var row = new TableRow(document);
    row.Cells.Add(GetSignerTableCell(document, leftBorder, "Штамп времени"));
    row.Cells.Add(GetSignerTableCell(document, rightBorder, signature.SigningTime));
    table.Rows.Add(row);
}

void AddSignatureTimeStampTime(DocumentCore document, Table table, Sign signature)
{
    var row = new TableRow(document);
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Left | MultipleBorderTypes.Bottom, "Пожпись заверена"));
    row.Cells.Add(GetSignerTableCell(document, MultipleBorderTypes.Right | MultipleBorderTypes.Bottom, signature.SignatureTimeStampTime));
    table.Rows.Add(row);
}

TableCell GetSignerTableCell(DocumentCore document, MultipleBorderTypes borderTypes, string text)
{
    var block = new Paragraph(document, new Run(document, text, new CharacterFormat
    {
        FontColor = Color.Blue,
        Size = DefaultFontSize
    }));

    var cell = new TableCell(document, block);

    cell.CellFormat.PreferredWidth = new TableWidth(50, TableWidthUnit.Percentage);
    cell.CellFormat.Borders.SetBorders(borderTypes, BorderStyle.Single, Color.Blue, 1);
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
    cell.CellFormat.Padding = new Padding(DefaultLeftPadding, DefaultPadding, DefaultPadding, DefaultPadding, LengthUnit.Millimeter);

    return cell;
}
