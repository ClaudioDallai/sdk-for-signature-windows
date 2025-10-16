using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using Shape = Microsoft.Office.Interop.Word.Shape;


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
        /// <param name="checkShapesText"> If true iterate through all shapes to check AlternativeTexts and retreive all (of the ones matching text) positions </param>
        /// <param name="outPath"> Output docx </param>
        public static void FillTemplate(out Dictionary<string, MarkerFoundLocationinfo> markersInfo, Dictionary<string, string> replacements, string inPath, string outPath, bool checkShapesText = false)
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
                    bool foundTextOccurence = false;

                    // Searh for texts markers
                    // First occurence to get position before it gets invalidated
                    foreach (Range storyRange in doc.StoryRanges)
                    {
                        Range searchRange = storyRange.Duplicate;
                        Find find = searchRange.Find;

                        find.ClearFormatting();
                        find.Text = pair.Key;
                        find.Forward = true;
                        find.Wrap = WdFindWrap.wdFindStop;

                        if (find.Execute())
                        {
                            if (!markersInfo.ContainsKey(pair.Key))
                            {
                                markersInfo.Add(pair.Key, new MarkerFoundLocationinfo
                                {
                                    Left = (float)searchRange.Information[WdInformation.wdHorizontalPositionRelativeToPage],
                                    Top = (float)searchRange.Information[WdInformation.wdVerticalPositionRelativeToPage],
                                    Page = (int)searchRange.Information[WdInformation.wdActiveEndPageNumber]
                                });

                                foundTextOccurence = true;
                                break;
                            }
                        }
                    }

                    if (foundTextOccurence)
                    {
                        foreach (Range storyRange in doc.StoryRanges)
                        {
                            Find replaceFind = storyRange.Find;
                            replaceFind.ClearFormatting();
                            replaceFind.Text = pair.Key;
                            replaceFind.Replacement.ClearFormatting();
                            replaceFind.Replacement.Text = pair.Value;
                            replaceFind.Forward = true;
                            replaceFind.Wrap = WdFindWrap.wdFindContinue;

                            replaceFind.Execute(Replace: WdReplace.wdReplaceAll);
                        }
                    }

                    // If flag is true, check for all AlternativeTexts inside Shapes, to find coordinates
                    // While this is a different operation, and a standalone method could exists, this is an heavy operation, so we call it here if necessary
                    if (!checkShapesText) continue;

                    // Searh for AlternativeTexts markers inside shapes
                    foreach (Shape shape in doc.Shapes)
                    {
                        if (shape.AlternativeText == pair.Key)
                        {
                            if (!markersInfo.ContainsKey(pair.Key))
                            {
                                Range anchorRange = shape.Anchor;
                                markersInfo.Add(pair.Key, new MarkerFoundLocationinfo
                                {
                                    Left = shape.Left,
                                    Top = shape.Top,
                                    Page = anchorRange.get_Information(WdInformation.wdActiveEndPageNumber)
                                });
                            }

                            shape.Delete();
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
