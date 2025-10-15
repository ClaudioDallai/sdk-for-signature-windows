using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;

namespace UtilityWordLib
{
    public class WordProcessor
    {
        public void FillTemplate(string inPath, string outPath, Dictionary<string, string> replacements)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;
            try
            {
                app.Visible = false;
                doc = app.Documents.Open(inPath);

                foreach (var pair in replacements)
                {
                    foreach (Range rng in doc.StoryRanges)
                    {
                        Find find = rng.Find;
                        find.ClearFormatting();
                        find.Text = pair.Key;
                        find.Replacement.ClearFormatting();
                        find.Replacement.Text = pair.Value;

                        find.Execute(Replace: WdReplace.wdReplaceAll);
                    }
                }

                doc.SaveAs2(outPath);
            }
            finally
            {
                doc?.Close(false);
                app.Quit(false);
            }
        }

        public void ConvertToPdf(string wordPath, string pdfPath)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;
            try
            {
                doc = app.Documents.Open(wordPath);
                doc.ExportAsFixedFormat(
                    pdfPath,
                    WdExportFormat.wdExportFormatPDF
                );
            }
            finally
            {
                doc?.Close(false);
                app.Quit(false);
            }
        }
    }
}
