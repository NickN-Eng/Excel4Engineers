using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel4Engineers
{
    /// <summary>
    /// The functions used in the [Formulae Text] group of the ribbon
    /// </summary>
    public static class MathFormatFunctions
    {
        #region Subscript superscript

        public const string SuperscriptRegexPattern = @"(?:\^)(.+?)(?: |_|$)";
        public const string SubscriptRegexPattern = @"(?:_)(.+?)(?: |^|$)";

        /// <summary>
        /// Applies the FormulaeText algorithm to selected cells
        /// </summary>
        /// <param name="rng">The selected range</param>
        /// <param name="subsuperscript">True if subscript/supercript mode has been enabled</param>
        /// <param name="symbols">True if symbol mode has been enabled</param>
        public static void FormatSubscriptSuperscript(Range rng, bool subsuperscript = true, bool symbols = true)
        {
            foreach (Range cel in rng.Cells)
            {
                var cellText = new XlCellText(cel);

                if (subsuperscript)
                {
                    foreach (Match m in Regex.Matches(cellText.Text, SuperscriptRegexPattern))
                    {
                        //Get the capture group (XX) with the text within the ^XX
                        var captureGroup = m.Groups[1];

                        //Ignore this if the text is already superscript
                        if (cellText.FontDatas[captureGroup.Index].Superscript) continue;

                        //Set the characters within the span as superscript and delete the ^
                        for (int i = captureGroup.Index; i < captureGroup.Index + captureGroup.Length; i++)
                        {
                            cellText.FontDatas[i].SetSuperscript();
                        }
                        cellText.FontDatas[captureGroup.Index - 1].SetDeleted();
                    }

                    foreach (Match m in Regex.Matches(cellText.Text, SubscriptRegexPattern))
                    {
                        //Get the capture group (XX) with the text within the _XX
                        var captureGroup = m.Groups[1];

                        //Ignore this if the text is already subscript
                        if (cellText.FontDatas[captureGroup.Index].Subscript) continue;

                        //Set the characters within the span as subscript and delete the _
                        for (int i = captureGroup.Index; i < captureGroup.Index + captureGroup.Length; i++)
                        {
                            cellText.FontDatas[i].SetSubscript();
                        }
                        cellText.FontDatas[captureGroup.Index - 1].SetDeleted();
                    }
                }

                if (symbols)
                {
                    foreach (var pair in SymbolLibrary)
                    {
                        foreach (var foundIndex in FindStringInText(cellText.Text, pair.Text))
                        {
                            if (cellText.CheckIfAnyDeleted(foundIndex, pair.Text.Length)) continue;

                            cellText.Delete(foundIndex, pair.Text.Length);
                            cellText.InsertText(foundIndex, pair.Symbol);
                        }
                    }
                }

                cellText.ApplyChanges();
            }
        }

        /// <summary>
        /// For the selected cells, reverts the application of the FormulaeText algorithm to the original cells.
        /// </summary>
        /// <param name="rng">The selected range</param>
        /// <param name="subsuperscript">True if subscript/supercript mode has been enabled</param>
        /// <param name="symbols">True if symbol mode has been enabled</param>
        public static void RevertSubscriptSuperscript(Range rng, bool subsuperscript = true, bool symbols = true)
        {
            foreach (Range cel in rng.Cells)
            {
                //string text = cel.Text;

                var cellText = new XlCellText(cel);


                if (symbols)
                {
                    foreach (var pair in SymbolLibrary)
                    {
                        foreach (var foundIndex in FindStringInText(cellText.Text, pair.Symbol))
                        {
                            if (cellText.CheckIfAnyDeleted(foundIndex, pair.Symbol.Length)) continue;

                            cellText.Delete(foundIndex, pair.Symbol.Length);
                            cellText.InsertText(foundIndex, pair.Text);
                        }
                    }
                }

                if (subsuperscript)
                {
                    bool isSuperscript = false;
                    bool isSubscript = false;
                    for (int i = 0; i < cellText.FontDatas.Length; i++)
                    {
                        var fd = cellText.FontDatas[i];

                        if (fd.Subscript)
                        {
                            cellText.FontDatas[i].SetStandardscript();
                            //fd.SetSuperscript();
                            //fd.SetStandardscript();
                            if (!isSubscript) 
                            {
                                isSubscript = true;
                                isSuperscript = false;
                                cellText.InsertText(i, "_");
                            }
                        }
                        else if (fd.Superscript)
                        {
                            cellText.FontDatas[i].SetStandardscript();
                            //fd.SetStandardscript();
                            fd.SetSubscript();
                            if (!isSuperscript)
                            {
                                isSubscript = false;
                                isSuperscript = true;
                                cellText.InsertText(i, "^");
                            }
                        }
                    }
                }


                cellText.ApplyChanges();
            }
        }



        #endregion

        #region Symbol replacement

        public class SymbolPair
        {
            public SymbolPair(string text, string symbol)
            {
                Text = text;
                Symbol = symbol;
            }

            public string Text { get; set; }   
            public string Symbol { get; set; }
        }

        public static IEnumerable<int> FindStringInText(string mainText, string substringToFind)
        {
            int startSearchPos = 0;
            while (true)
            {
                int foundIndex = mainText.IndexOf(substringToFind, startSearchPos);

                if (foundIndex == -1) break;

                yield return foundIndex;
                startSearchPos = foundIndex + substringToFind.Length;
            } 
        }

        public static List<SymbolPair> SymbolLibrary = new List<SymbolPair>
        {
            new SymbolPair ( @"\Alpha", "Α" ),
            new SymbolPair ( @"\Beta", "Β" ),
            new SymbolPair ( @"\Gamma", "Γ" ),
            new SymbolPair ( @"\Delta", "Δ" ),
            new SymbolPair ( @"\Epsilon", "Ε" ),
            new SymbolPair ( @"\Zeta", "Ζ" ),
            new SymbolPair ( @"\Eta", "Η" ),
            new SymbolPair ( @"\Theta", "Θ" ),
            new SymbolPair ( @"\Iota", "Ι" ),
            new SymbolPair ( @"\Kappa", "Κ" ),
            new SymbolPair ( @"\Lambda", "Λ" ),
            new SymbolPair ( @"\Mu", "Μ" ),
            new SymbolPair ( @"\Nu", "Ν" ),
            new SymbolPair ( @"\Xi", "Ξ" ),
            new SymbolPair ( @"\Omicron", "Ο" ),
            new SymbolPair ( @"\Pi", "Π" ),
            new SymbolPair ( @"\Rho", "Ρ" ),
            new SymbolPair ( @"\Sigma", "Σ" ),
            new SymbolPair ( @"\Tau", "Τ" ),
            new SymbolPair ( @"\Upsilon", "Υ" ),
            new SymbolPair ( @"\Phi", "Φ" ),
            new SymbolPair ( @"\Chi", "Χ" ),
            new SymbolPair ( @"\Psi", "Ψ" ),
            new SymbolPair ( @"\Omega", "Ω" ),
            new SymbolPair ( @"\alpha", "α" ),
            new SymbolPair ( @"\beta", "β" ),
            new SymbolPair ( @"\gamma", "γ" ),
            new SymbolPair ( @"\delta", "δ" ),
            new SymbolPair ( @"\epsilon", "ε" ),
            new SymbolPair ( @"\zeta", "ζ" ),
            new SymbolPair ( @"\eta", "η" ),
            new SymbolPair ( @"\theta", "θ" ),
            new SymbolPair ( @"\iota", "ι" ),
            new SymbolPair ( @"\kappa", "κ" ),
            new SymbolPair ( @"\lambda", "λ" ),
            new SymbolPair ( @"\mu", "μ" ),
            new SymbolPair ( @"\nu", "ν" ),
            new SymbolPair ( @"\xi", "ξ" ),
            new SymbolPair ( @"\omicron", "ο" ),
            new SymbolPair ( @"\pi", "π" ),
            new SymbolPair ( @"\rho", "ρ" ),
            new SymbolPair ( @"\sigma", "σ" ),
            new SymbolPair ( @"\tau", "τ" ),
            new SymbolPair ( @"\upsilon", "υ" ),
            new SymbolPair ( @"\phi", "φ" ),
            new SymbolPair ( @"\chi", "χ" ),
            new SymbolPair ( @"\psi", "ψ" ),
            new SymbolPair ( @"\omega", "ω" )
        };

        #endregion


        #region XlCellText

        /// <summary>
        /// A temporary class holding sub & superscript data
        /// You can apply making changes to the data in this class, and apply these changes back to the cell font format
        /// This class is all about applying the minimal changes to the text within excel.
        /// </summary>
        public class XlCellText
        {
            public string Text { get; set; }
            public XlFontData[] FontDatas { get; set; }

            private Range Cell { get; set; }

            //private SortedList<int, string> TextToInsert { get; set; } = new SortedList<int, string>();
            private List<KeyValuePair<int, string>> TextToInserts { get; set; } = new List<KeyValuePair<int, string>>();

            public XlCellText(Range cell)
            {
                Cell = cell;
                //Getting text may fail if the cell is a number format,
                //so catch that case and change to text ("@") if it happens
                try
                {
                    Text = cell.Characters.Text;
                }
                catch (Exception)
                {
                    cell.NumberFormat = "@";
                    Text = cell.Characters.Text;
                }
                var length = Text.Length;

                FontDatas = new XlFontData[length];

                //Get the superscript data if that happens
                for (int i = 1; i <= length; i++)
                {
                    var font = cell.Characters[i, 1].Font;

                    XlFontData newFontData;
                    //Get subscript superscript data
                    if (font.Superscript) newFontData = XlFontData.CreateSuperscript();
                    else if (font.Subscript) newFontData = XlFontData.CreateSubscript();
                    else newFontData = new XlFontData();

                    FontDatas[i-1] = newFontData;
                }
                    
            }

            /// <summary>
            /// Mark text to be inserted.
            /// Inserted text will follow formatting of the insertion position
            /// </summary>
            public void InsertText(int insertPosition, string text) 
            { 
                //TextToInsert[insertPosition] = text; 
                TextToInserts.Add(new KeyValuePair<int, string>(insertPosition, text));
            }

            /// <summary>
            /// Find all the texts that have been inserted at a certain position
            /// </summary>
            /// <param name="insertPosition"></param>
            /// <returns></returns>
            public IEnumerable<string> GetInsertionTexts(int insertPosition)
            {
                foreach (var kvp in TextToInserts)
                {
                    if(kvp.Key == insertPosition) yield return kvp.Value;
                }
            }

            /// <summary>
            /// Applies changes (i.e. subscript/superscript/inserted/deleted characters)
            /// to the selected cell
            /// </summary>
            public void ApplyChanges()
            {
                //Apply subscript superscript changes (excluding newly added 
                for (int i = 0; i < FontDatas.Length; i++)
                {
                    var iFont = FontDatas[i];
                    if (!iFont.SSSScriptChanged || iFont.Deleted) continue;

                    if(iFont.Superscript)
                        Cell.Characters[i+1,1].Font.Superscript = true;
                    else if (iFont.Subscript)
                        Cell.Characters[i+1,1].Font.Subscript = true;
                    else
                    {
                        Cell.Characters[i + 1, 1].Font.Subscript = false;
                        Cell.Characters[i + 1, 1].Font.Superscript = false;
                    }
                }

                //Apply deletions and insertions
                //Go in reverse to avoid affecting indices
                for (int i = FontDatas.Length - 1; i >= 0; i--)
                {
                    var iFont = FontDatas[i];
                    var iChar = Cell.Text[i];

                    if (iFont.Deleted)
                        Cell.Characters[i+1,1].Delete();

                    foreach (var textToInsert in GetInsertionTexts(i))
                    {
                        Cell.Characters[i + 1, 0].Insert(textToInsert);
                         //If this char is deleted, check the previous char to see whether it should be subscript or superscript
                        if (iFont.Deleted)
                        {
                            if (i <= 0)
                            {
                                //Do nothing if i = 0, since we cannot take i-1
                            }
                            else if (Text[i - 1] == '_')
                            {
                                Cell.Characters[i + 1, textToInsert.Length].Font.Subscript = true;
                            }
                            else if (Text[i - 1] == '^')
                            {
                                Cell.Characters[i + 1, textToInsert.Length].Font.Superscript = true;
                            }
                            else if (FontDatas[i - 1].Subscript)
                            {
                                Cell.Characters[i + 1, textToInsert.Length].Font.Subscript = true;
                            }
                            else if (FontDatas[i - 1].Superscript)
                            {
                                Cell.Characters[i + 1, textToInsert.Length].Font.Superscript = true;
                            }
                        }
                    }
                }
            }

            public bool CheckIfAnyDeleted(int startIndex, int length)
            {
                for (int i = 0; i < length; i++)
                {
                    if (FontDatas[i+startIndex].Deleted) return true;
                }
                return false;
            }

            public void Delete(int startIndex, int length)
            {
                for (int i = 0; i < length; i++)
                {
                    FontDatas[i + startIndex].SetDeleted();
                }
            }
        }

        /// <summary>
        /// Represents the Superscript/Subscript/Deleted state of a character
        /// This object is meant to be created witht the initial state of the character,
        /// then SetSuperscript/SetSubscript/SetDeleted which tracks changes.
        /// Then applied to the original text and discarded. Cannot revert back to unchanged state.
        /// </summary>
        public struct XlFontData
        {
            private bool _Deleted;
            private bool _Superscript;
            private bool _Subscript;
            private bool _SSSScriptChanged;

            public bool Deleted => _Deleted;
            public bool Superscript => _Superscript;
            public bool Subscript => _Subscript;
            public bool SSSScriptChanged => _SSSScriptChanged;

            public void SetSubscript()
            {
                _Subscript = true;
                _SSSScriptChanged = true;
            }

            public void SetSuperscript()
            {
                _Superscript = true;
                _SSSScriptChanged = true;
            }

            public void SetStandardscript()
            {
                _Superscript = false;
                _Subscript = false;
                _SSSScriptChanged = true;
            }

            public void SetDeleted()
            {
                _Deleted = true;
            }

            public static XlFontData CreateSuperscript()
            {
                return new XlFontData()
                {
                    _Superscript = true
                };
            }

            public static XlFontData CreateSubscript()
            {
                return new XlFontData()
                {
                    _Subscript = true
                };
            }
        }

        #endregion

    }
}
