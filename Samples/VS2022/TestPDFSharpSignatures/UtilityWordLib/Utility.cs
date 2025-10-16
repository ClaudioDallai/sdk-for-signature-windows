using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;


namespace UtilityWordLib
{
    public struct MarkerFoundLocationinfo
    {
        public float Left;
        public float Top;
        public int Page;
    }

    public static class WordProcessor
    {

        /// <summary>
        /// Fill a docx using text markers
        /// </summary>
        /// <param name="replacements"> Text markers replacements </param>
        /// <param name="inPath"> Docx template to fill </param>
        /// <param name="outPath"> Output docx </param>
        public static void FillTemplate(Dictionary<string, string> replacements, string inPath, string outPath)
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


        /// <summary>
        /// Fill a docx using text markers
        /// </summary>
        /// <param name="markersInfo"> Dict containing markers position datas. Beware if multiple same markers, only the first one will be returned </param>
        /// <param name="replacements"> Text markers replacements </param>
        /// <param name="inPath"> Docx template to fill </param>
        /// <param name="outPath"> Output docx </param>
        public static void FillTemplate(out Dictionary<string, MarkerFoundLocationinfo> markersInfo, Dictionary<string, string> replacements, string inPath, string outPath)
        {
            markersInfo = new Dictionary<string, MarkerFoundLocationinfo>();

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

                        if (find.Execute(Replace: WdReplace.wdReplaceAll))
                        {

                            // Only new markers for the first time
                            if (!markersInfo.ContainsKey(pair.Key))
                            {
                                markersInfo.Add(pair.Key, new MarkerFoundLocationinfo
                                {
                                    Left = (float)rng.Information[WdInformation.wdHorizontalPositionRelativeToPage],
                                    Top = (float)rng.Information[WdInformation.wdVerticalPositionRelativeToPage],
                                    Page = (int)rng.Information[WdInformation.wdActiveEndPageNumber]
                                });
                            }

                        }
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

        public static void ConvertToPdf(string wordPath, string pdfPath)
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
