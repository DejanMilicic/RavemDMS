namespace Dms;

using Models;
using Raven.Client.Documents.Indexes;

public class Search : AbstractIndexCreationTask<Doc>
{
    public override string IndexName => "Documents/Search";

    public class Entry
    {
        public string[] Query { get; set; }
    }

    public Search()
    {
        Map = docs => from doc in docs
                let attachments = AttachmentsFor(doc)
                from att in attachments
                    let stream = LoadAttachment(doc, att.Name).GetContentAsStream()
                    select new
                    {
                        Query = ExtractText.FromAttachment(att.Name, stream)
                    };

        Index("Query", FieldIndexing.Search);

        AdditionalAssemblies = new HashSet<AdditionalAssembly>()
        {
            AdditionalAssembly.FromNuGet("DocumentFormat.OpenXml", "2.14.0"),
            AdditionalAssembly.FromNuGet("iTextSharp", "5.5.13.2")
        };

        AdditionalSources = new Dictionary<string, string>
        {
            ["ExtractText.cs"] =
                File.ReadAllText(Path.Combine(new[] { AppContext.BaseDirectory, "..", "..", "..", "ExtractText.cs" }))
        };
    }
}
