using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace DocumentProject
{
    public partial class ThisDocument
    {
        private void AddBibliography()
        {
            Word.WdLanguageID lcid = Word.WdLanguageID.wdEnglishUS;
            string guid = System.Guid.NewGuid().ToString();
            string src =
                "<b:Source><b:Tag>Jam08</b:Tag><b:SourceType>Book</b:SourceType>"
                + "<b:Guid>" + guid + "</b:Guid><b:LCID>0</b:LCID><b:Author>"
                + "<b:Author><b:NameList><b:Person><b:Last>Persse</b:Last>"
                + "<b:First>James</b:First></b:Person></b:NameList></b:Author>"
                + "</b:Author><b:Title>Hollywood Secrets of Project Management "
                + "Success</b:Title><b:Year>2008</b:Year><b:City>Redmond</b:City>"
                + "<b:Publisher>Microsoft Press</b:Publisher></b:Source>";

            this.Bibliography.Sources.Add(src);

            this.Bibliography.BibliographyStyle = "APA";
            this.Paragraphs.Last.Range.InsertParagraphAfter();

            object fieldType = Word.WdFieldType.wdFieldBibliography;
            object missing = System.Reflection.Missing.Value;

            this.Fields.Add(
                this.Paragraphs.Last.Range,
                ref fieldType,
                ref missing,
                ref missing);
        }
    }

}
