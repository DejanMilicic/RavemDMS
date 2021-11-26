using Dms;
using Dms.Models;
using Raven.Client.Documents;
using Raven.Client.Documents.Session;

var session = DocumentStoreHolder.Store.OpenSession();

//Seed();

string term = "ferrari";

List<Doc> results = session.Query<Search.Entry, Search>()
        .Search(x => x.Query, "ferrari")
        .Statistics(out QueryStatistics stats)
        //.Skip(0)
        //.Take(1)
        .ProjectInto<Doc>()
        .ToList();

Console.WriteLine($"Searching for '{term}'");
Console.WriteLine($"Elapsed time: {stats.DurationInMs}ms");
Console.WriteLine($"Total results: {stats.TotalResults}");
Console.WriteLine("=== Results:");

foreach (Doc doc in results)
    Console.WriteLine($"Doc.Id: {doc.Id} \t Doc.Name: {doc.Name}");

void Seed()
{
    Doc wordDoc = new Doc
    {
        Name = "word"
    };
    session.Store(wordDoc);
    using var wordStream = new FileStream(@"..\..\..\..\..\..\docs\Lorem.docx", FileMode.Open);
    session.Advanced.Attachments.Store(wordDoc.Id, "Loren.docx", wordStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    Doc excelDoc = new Doc
    {
        Name = "excel"
    };
    session.Store(excelDoc);
    using var excelStream = new FileStream(@"..\..\..\..\..\..\docs\Excel.xlsx", FileMode.Open);
    session.Advanced.Attachments.Store(excelDoc.Id, "Excel.xlsx", excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    Doc pdfDoc = new Doc
    {
        Name = "pdf"
    };
    session.Store(pdfDoc);
    using var pdfStream = new FileStream(@"..\..\..\..\..\..\docs\Pdf.pdf", FileMode.Open);
    session.Advanced.Attachments.Store(pdfDoc.Id, "Pdf.pdf", pdfStream, "application/pdf");

    Doc imageDoc = new Doc
    {
        Name = "image"
    };
    session.Store(imageDoc);
    using var imageStream = new FileStream(@"..\..\..\..\..\..\docs\image.jpg", FileMode.Open);
    session.Advanced.Attachments.Store(imageDoc.Id, "image.jpg", imageStream, "image/jpeg");

    session.SaveChanges();




}