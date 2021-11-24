using Dms;
using Dms.Models;

var session = DocumentStoreHolder.Store.OpenSession();

Seed();

void Seed()
{
    Doc wordDoc = new Doc();
    session.Store(wordDoc);
    using var stream = new FileStream(@"..\..\..\..\..\..\docs\Lorem.docx", FileMode.Open);
    session.Advanced.Attachments.Store(wordDoc.Id, "Loren.docx", stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    Doc excelDoc = new Doc();
    session.Store(excelDoc);
    using var excelStream = new FileStream(@"..\..\..\..\..\..\docs\Excel.xlsx", FileMode.Open);
    session.Advanced.Attachments.Store(excelDoc.Id, "Excel.xlsx", excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    Doc pdfDoc = new Doc();
    session.Store(pdfDoc);
    using var pdfStream = new FileStream(@"..\..\..\..\..\..\docs\Pdf.pdf", FileMode.Open);
    session.Advanced.Attachments.Store(excelDoc.Id, "Pdf.pdf", pdfStream, "application/pdf");

    session.SaveChanges();
}