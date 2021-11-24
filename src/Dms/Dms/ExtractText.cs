namespace Dms;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

public static class ExtractText
{
    public static IEnumerable<string> FromAttachment(string fileName, Stream stream)
    {
        if (fileName.ToLower().EndsWith(".docx"))
        {
            return GetWordText(stream);
        }
        else
        {
            return new string[]
            {
                fileName
            };
        }
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
}

