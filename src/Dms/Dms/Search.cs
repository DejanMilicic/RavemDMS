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
            select new
            {
                Query = ExtractText.FromAttachment(att.Name)
            };

        Index("Query", FieldIndexing.Search);

        AdditionalSources = new Dictionary<string, string>
        {
            ["ExtractText.cs"] = @"
using System;
using System.Collections.Generic;

public static class ExtractText
{
    public static string[] FromAttachment(string fileName)
    {
        return new string[]
        {
            fileName
        };
    }
}


"
        };
    }
}
