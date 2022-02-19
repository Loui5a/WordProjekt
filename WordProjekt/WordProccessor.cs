using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProjekt
{
    internal class WordProccessor
    {
        private const string FileLocation = "testdoc.docx";

        public bool MakeEnrichedCopy(enrichmentData data)
        {
            var file = File.OpenRead(FileLocation);
            var memoryStream = new MemoryStream();
            file.CopyTo(memoryStream);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                TextReplacer.SearchAndReplace(doc, "#REPLACEME#", "replacementtext", false);

                TextReplacer.SearchAndReplace(doc, "#REPLACEME2#", "pizza", false);

                TextReplacer.SearchAndReplace(doc, "#REPLACEME3#", "burger", false);
                doc.SaveAs("resultdoc.docx");
                return true;

            }
        }
    }
}
