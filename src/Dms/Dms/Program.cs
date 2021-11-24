using Dms;
using Dms.Models;

var session = DocumentStoreHolder.Store.OpenSession();

//Seed();

void Seed()
{
    Doc doc = new Doc();
    session.Store(doc);
    
    using var stream = new FileStream(@"..\..\..\..\..\..\docs\Lorem.docx", FileMode.Open);
    session.Advanced.Attachments.Store(doc.Id, "Loren.docx", stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    session.SaveChanges();
}