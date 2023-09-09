using System;
using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig.AcroForms;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Geometry;

namespace Excel4Engineers
{
    /// <summary>
    /// A reader using PdfPig which reads annotations on the TemplatePdf and uses the consituent annotations to 
    /// scan for pdf text in ScanPdf(pdf)
    /// </summary>
    public class PdfReader
    {
        /// <summary>
        /// Error messages created when parsing the template pdf
        /// </summary>
        public List<string> TemplateMessages { get; private set; } = new List<string>();

        /// <summary>
        /// Warning messages created during running of ScanPdf()
        /// </summary>
        public List<string> Warnings { get; private set; } = new List<string>();

        /// <summary>
        /// The list of search boxes
        /// </summary>
        public List<SearchBox> SearchBoxes { get; private set; }


        /// <summary>
        /// Set a template file with annotations showing the search regions.
        /// The filepath must be valid (i.e. prior checking should have taken place).
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool SetTemplate(string filename)
        {
            TemplateMessages.Clear();

            using (PdfDocument document = PdfDocument.Open(filename))
            {
                if (document.NumberOfPages == 0)
                {
                    TemplateMessages.Add("Template Pdf has no pages.");
                    return false;
                }

                if (document.NumberOfPages > 1)
                {
                    TemplateMessages.Add("Pdf has more than 1 page, only the first page will be used.");
                }

                UglyToad.PdfPig.Content.Page firstPage = document.GetPages().FirstOrDefault();
                if (firstPage == null)
                    throw new Exception("Pdf has more than 1 page, but first page is null? Investigate...");

                SearchBoxes = SearchBox.GetSearchBoxes(firstPage);

                var words = firstPage.GetWords().ToList();
                var TextList = SearchBoxes.Select(b => b.GetText(words)).ToList();

                TemplateMessages.Add($"Found {SearchBoxes.Count}No. search regions.:");

                foreach (var box in SearchBoxes)
                {
                    var message = $"> Region tagged \"{box.Name}\" @[{box.Box.Centroid.X},{box.Box.Centroid.Y}] with example text: {box.GetText(words)}";
                    TemplateMessages.Add(message);
                }

                return true;
            }
        }

        /// <summary>
        /// Scan a pdf, using the search boxes defined from the pdf template
        /// </summary>
        /// <param name="pdfPath">The pdf to scan</param>
        /// <param name="foundText">The text found in each of the search boxes
        /// in the same order of the SearchBoxes</param>
        /// <returns>True if sucessful. This may fail if the pdf has no pages.</returns>
        public bool ScanPdf(string pdfPath, out List<string> foundText)
        {
            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                if (document.NumberOfPages == 0)
                {
                    Warnings.Add("Pdf has no pages.");
                    foundText = null;
                    return false;
                }

                if (document.NumberOfPages > 1)
                {
                    Warnings.Add("Pdf has more than 1 page, only the first page will be used.");
                }

                var firstPage = document.GetPages().FirstOrDefault();
                if (firstPage == null)
                    throw new Exception("Pdf has more than 1 page, but first page is null? Investigate...");

                var words = firstPage.GetWords().ToList();

                foundText = SearchBoxes.Select(b => b.GetText(words)).ToList();

                return true;
            }
        }

        /// <summary>
        /// Creates a printable string with all warnings from ScanPdf()
        /// </summary>
        /// <param name="preamble">Text which will printed at the beginning of the string, seperated by a newline</param>
        /// <returns>Error message</returns>
        public string PrintWarnings(string preamble)
        {
            return preamble + Environment.NewLine + string.Join(Environment.NewLine, Warnings.ToArray());
        }

        /// <summary>
        /// Show the error messages from ScanPdf() in a Messagebox
        /// </summary>
        /// <param name="preamble">Text which will printed at the beginning of the string, seperated by a newline</param>
        public void ShowMessageboxWithWarnings(string preamble)
        {
            System.Windows.Forms.MessageBox.Show(PrintWarnings(preamble));
        }

        /// <summary>
        /// Creates a printable string with all warnings from SetTemplate()
        /// </summary>
        /// <param name="preamble">Text which will printed at the beginning of the string, seperated by a newline</param>
        /// <returns>Error message</returns>
        public string PrintTemplateMessages(string preamble)
        {
            return preamble + Environment.NewLine + string.Join(Environment.NewLine, TemplateMessages.ToArray());
        }

        /// <summary>
        /// Show the error messages from SetTemplate() in a Messagebox
        /// </summary>
        /// <param name="preamble">Text which will printed at the beginning of the string, seperated by a newline</param>
        public void ShowMessageboxTemplateMessages(string preamble)
        {
            System.Windows.Forms.MessageBox.Show(PrintTemplateMessages(preamble));
        }

        /// <summary>
        /// Clear warnings from ScanPdf()
        /// </summary>
        public void ClearWarnings() => Warnings.Clear();


        /// <summary>
        /// A class contain a search region, which can find text within provided pdfs
        /// </summary>
        public class SearchBox
        {
            public string Name;

            public PdfRectangle Box;

            //public static string[] Names = new string[] { "TITLE", "DRAWING NUMBER", "REV" };

            /// <summary>
            /// Gets all the search boxes from a template pdf
            /// Basically reads all the annotations and uses thier rectnagular regions as the search box region.
            /// </summary>
            /// <param name="page"></param>
            /// <returns></returns>
            public static List<SearchBox> GetSearchBoxes(Page page)
            {
                var annots = page.ExperimentalAccess.GetAnnotations();
                var searchBoxes = new List<SearchBox>();

                foreach (var annot in annots)
                {
                    var text = annot.Content;

                    //Implentation 1: Use all annotations as search regions
                    searchBoxes.Add(new SearchBox()
                    {
                        Name = text,
                        Box = annot.Rectangle
                    });

                    //Implentation 2: Only use annotations which match the "Names" list
                    //if (Names.Contains(text))
                    //{
                    //    searchBoxes.Add(new SearchBox()
                    //    {
                    //        Name = text,
                    //        Box = annot.Rectangle
                    //    });
                    //}
                }

                return searchBoxes;
            }

            /// <summary>
            /// Filter all words within the page using this SearchBox
            /// </summary>
            /// <param name="wordsOnPage">Full list of words on the page</param>
            /// <returns>Words that are within this search box</returns>
            public List<Word> GetWordsWithinBox(List<Word> wordsOnPage)
            {
                var result = new List<Word>();
                foreach (var w in wordsOnPage)
                {
                    if (Box.Contains(w.BoundingBox.Centroid))
                    {
                        result.Add(w);
                    }
                }
                return result;
            }

            /// <summary>
            /// Gets text within the provided words
            /// </summary>
            /// <param name="wordsOnPage"></param>
            public string GetText(List<Word> wordsOnPage)
            {
                var wordList = GetWordsWithinBox(wordsOnPage);

                if (wordList.Count == 0) return null;
                if (wordList.Count == 1) return wordList[0].Text;

                var blocks = DefaultPageSegmenter.Instance.GetBlocks(wordList);

                return string.Join(" ", blocks.Select(b => b.Text)).Replace(Environment.NewLine, " ").Replace('\n', ' ');
            }
        }
    }

}
