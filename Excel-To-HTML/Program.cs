using Excel_To_HTML;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

const string ExcelToPdfSerialNumber = "...";
const string DocumentNetSerialNumber = "...";
const string FileName = "SampleDoc";
const int DefaultMargin = 0;

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
    excelToPdf.PageStyle.PageOrientation.Auto();
    excelToPdf.PageStyle.PageSize.Auto();

    var pdf = excelToPdf.ConvertBytes(excel);

    await File.WriteAllBytesAsync($".\\{FileName}.pdf", pdf);

    return pdf;
}

async Task ConvertToHtmlAsync(byte[] pdf)
{
    var document = LoadDocument(pdf);
    var builder = new DocumentBuilder(document);
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
    var sign = new Sign(certificate, DateTime.UtcNow, DateTime.UtcNow);

    AddSignatureStamp(builder, sign);

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

void AddSignatureStamp(DocumentBuilder builder, Sign sign)
{
    builder.ParagraphFormat.KeepLinesTogether = true;
    builder.ParagraphFormat.KeepWithNext = true;
    builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.None, Color.Blue, 1);
    builder.Writeln();

    builder.StartTable();

    builder.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(6.5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
    builder.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Blue, 1);
    builder.TableFormat.Alignment = HorizontalAlignment.Center;

    builder.CellFormat.Padding = new Padding(0.5, 0.2, LengthUnit.Centimeter);
    builder.CellFormat.VerticalAlignment = VerticalAlignment.Center;
    builder.ParagraphFormat.Alignment = HorizontalAlignment.Left;

    builder.RowFormat.Height = new TableRowHeight(5f, HeightRule.Exact);
    builder.InsertCell();
    builder.CellFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(1.8, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
    builder.InsertCell();
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.CharacterFormat.FontColor = Color.Blue;
    builder.CharacterFormat.Size = 8;
    builder.CharacterFormat.FontName = "Arial";
    builder.Write("SignedWith");
    builder.InsertCell();
    builder.Write($"{sign.OrganizationName} {sign.Employee}");
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.Write("Serial");
    builder.InsertCell();
    builder.Write(sign.SerialNumber);
    builder.EndRow();

    builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
    builder.InsertCell();
    builder.Write("ValidationPeriod");
    builder.InsertCell();
    builder.Write(sign.ValidityPeriod);
    builder.EndRow();

    if (!string.IsNullOrEmpty(sign.SignatureTimeStampTime) || !string.IsNullOrEmpty(sign.SigningTime))
    {
        builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Dashed, Color.Blue, 1);
        builder.RowFormat.Height = new TableRowHeight(5f, HeightRule.Exact);
        builder.InsertCell();
        builder.InsertCell();
        builder.EndRow();

        builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.None, Color.Blue, 1);

        if (!string.IsNullOrEmpty(sign.SigningTime))
        {
            builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
            builder.InsertCell();
            builder.Write("SignationDate");
            builder.InsertCell();
            builder.Write(sign.SigningTime);
            builder.EndRow();
        }

        if (!string.IsNullOrEmpty(sign.SignatureTimeStampTime))
        {
            builder.RowFormat.Height = new TableRowHeight(18f, HeightRule.Exact);
            builder.InsertCell();
            builder.Write("SignatureTimeStamp");
            builder.InsertCell();
            builder.Write(sign.SignatureTimeStampTime);
            builder.EndRow();
        }
    }

    builder.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.Single, Color.Blue, 1);
    builder.RowFormat.Height = new TableRowHeight(1f, HeightRule.Exact);
    builder.InsertCell();
    builder.InsertCell();
    builder.EndRow();

    builder.EndTable();
}
