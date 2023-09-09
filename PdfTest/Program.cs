//using iTextSharp.text.pdf.parser;
//using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig.AcroForms;
using UglyToad.PdfPig.Annotations;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;

namespace PdfTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var filename = @"C:\Users\nicho\OneDrive\Coding\Pdf scanning\TestDrawings\RambollDrawings\1620011788-RAM-AB-00-DR-S-00100_TEMPLATE.pdf";
            using (PdfDocument document = PdfDocument.Open(filename))
            {
                // Get the title from the document metadata.
                Console.WriteLine(document.Information.Title);

                var struc = document.Structure;
                var bResult = document.TryGetBookmarks(out Bookmarks book);
                var xResult = document.TryGetXmpMetadata(out XmpMetadata metadata);
                var fResult = document.TryGetForm(out AcroForm form);

                foreach (Page page in document.GetPages())
                {
                    var words = page.GetWords().ToList();


                    var boxes = SearchBox.GetSearchBoxes(page);


                    //foreach (var word in words)
                    //{
                    //    var wordCentre = word.BoundingBox.Centroid;
                    //    foreach (var box in boxes)
                    //    {
                    //        box.Box.Contains(wordCentre);
                    //    }
                    //}

                    var wordList = boxes.Select(b => b.GetWords(words)).ToList();
                    var TextList = boxes.Select(b => b.GetText(words)).ToList();

                }

                
            }

            //AnnotReader.ReadAnnotations(filename);
            //StringBuilder text = new StringBuilder();

            //if (File.Exists(filename))
            //{
            //    PdfReader pdfReader = new PdfReader(filename);

            //    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            //    {
            //        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            //        ITextExtractionStrategy strategy2 = new LocationTextExtractionStrategy();
            //        string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

            //        currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
            //        text.Append(currentText);
            //    }
            //    pdfReader.Close();
            //}
            //var text = text.ToString();
        }


    }

    public class SearchBox
    {
        public string Name;

        public PdfRectangle Box;

        public static string[] Names = new string[] { "TITLE", "DRAWING NUMBER", "REV" };

        public static List<SearchBox> GetSearchBoxes(Page page)
        {
            var annots = page.ExperimentalAccess.GetAnnotations();
            var searchBoxes = new List<SearchBox>();

            foreach (var annot in annots)
            {
                var text = annot.Content;
                if (Names.Contains(text))
                {
                    searchBoxes.Add(new SearchBox()
                    {
                        Name = text,
                        Box = annot.Rectangle
                    });
                }
            }

            return searchBoxes;
        }

        public List<Word> GetWords(List<Word> words)
        {
            var result = new List<Word>();
            foreach (var w in words)
            {
                if(Box.Contains(w.BoundingBox.Centroid))
                {
                    result.Add(w);
                }
            }
            return result;

            var blocks = DefaultPageSegmenter.Instance.GetBlocks(words);
        }

        public string GetText(List<Word> words)
        {
            var wordList = GetWords(words);
            
            if(wordList.Count == 0) return null;
            if(wordList.Count == 1) return wordList[0].Text;

            var blocks = DefaultPageSegmenter.Instance.GetBlocks(wordList);

            return string.Join(" ", blocks.Select(b => b.Text)).Replace(Environment.NewLine, " ");
        }
    }
}
