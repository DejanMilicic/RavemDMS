using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

public static class ExtractText
{
    public static IEnumerable<string> FromAttachment(string fileName, Stream stream)
    {
        if (fileName.ToLower().EndsWith(".docx"))
        {
            return GetWordText(stream);
        }
        else if (fileName.ToLower().EndsWith(".xlsx"))
        {
            return GetExcelText(stream);
        }
        else if (fileName.ToLower().EndsWith(".pdf"))
        {
            return GetPdfText(stream);
        }
        else
        {
            return new string[]
            {
                fileName
            };
        }
    }

    // nuget: iTextSharp 5.5.13.2
    private static IEnumerable<string> GetPdfText(Stream stream)
    {
        List<string> content = new List<string>();
        
        PdfReader reader = new PdfReader(stream);
        for (int page = 1; page <= reader.NumberOfPages; page++)
            content.Add(PdfTextExtractor.GetTextFromPage(reader, page));
        reader.Close();
        
        return content;
    }

    // nuget: DocumentsFormat.OpenXml
    public static IEnumerable<string> GetWordText(Stream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        foreach (var element in doc.MainDocumentPart.Document.Body)
        {
            if (element is Paragraph p)
            {
                yield return p.InnerText;
            }
        }

        var comments = doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
        if (comments == null)
            yield break;
        foreach (var element in comments)
        {
            yield return element.InnerText;
        }
    }

    public static IEnumerable<string> GetExcelText(Stream stream)
    {
        using var doc = SpreadsheetDocument.Open(stream, false);
        Func<OpenXmlElement, string> selector = x => x.InnerText;

        IEnumerable<SharedStringTablePart> sharedTableParts = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>();

        string[] sst = sharedTableParts
            .First()
            .SharedStringTable.ChildElements.Select(selector)
            .ToArray();

        foreach (var sheet in doc.WorkbookPart.Workbook.Descendants<Sheet>())
        {
            var part = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            foreach (var cell in part.Worksheet.Descendants<Cell>())
            {
                switch (cell.DataType?.Value)
                {
                    case CellValues.Boolean:
                        yield return cell.InnerText == "0" ? "false" : "true";
                        break;
                    case CellValues.SharedString:
                        yield return sst[int.Parse(cell.InnerText)];
                        break;
                    case CellValues.Date:
                        yield return cell.InnerText;
                        break;
                }
            }
        }
    }
}

