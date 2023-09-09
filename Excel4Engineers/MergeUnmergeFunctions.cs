using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel4Engineers
{
    /// <summary>
    /// The functions used in the [Merge Unmerge] group of the ribbon
    /// </summary>
    public static class MergeUnmergeFunctions
    {
        /// <summary>
        /// Merges every cell in the range.
        /// Cells in the same row are joined, cells in subsequent rows are joined on the next line
        /// </summary>
        /// <param name="selectionRange"></param>
        public static void MergeWithoutDelete(Range selectionRange)
        {
            if (selectionRange == null | selectionRange.Count == 0) return;

            var cellData = new List<string>();
            foreach (Range a in selectionRange.Areas)
            {
                var sb = new StringBuilder();
                Range firstCell = a.FirstCell();
                int rowNo = firstCell.Row;
                foreach (Range cel in a.Cells)
                {
                    if (rowNo != cel.Row) sb.AppendLine(); //Append a newline when the row changes
                    if(cel.Value2 != null)
                        sb.Append(cel.Value2);

                    rowNo = cel.Row;
                }
                a.Clear();
                a.Merge();
                firstCell.Value2 = sb.ToString();
            }

        }

        /// <summary>
        /// Merges every cell in the selection accross the rows
        /// Text in the cells being merged are joined with a space " "
        /// </summary>
        public static void MergeAcrossJoin(Range selectionRange)
        {
            if (selectionRange == null | selectionRange.Count == 0) return;

            var cellData = new List<string>();
            foreach (Range a in selectionRange.Areas)
            {
                foreach (Range row in a.Rows)
                {
                    var rowFirstCel = row.FirstCell();
                    var joinedTxt = string.Join(" ", row.GetValuesAsString().Where(str => !string.IsNullOrEmpty(str)));
                    row.Clear();
                    row.Merge();
                    rowFirstCel.Value2 = joinedTxt;
                }
            }
        }

        /// <summary>
        /// Unmerges any grouped cells in the selection range provided.
        /// </summary>
        /// <param name="range"></param>
        public static void Unmerge(Range range, UnmergeMode mode)
        {
            if (range == null | range.Count == 0) return;

            var mergedAreas = new Dictionary<string, Range>();
            foreach (Range cel in range.Cells)
            {
                
                //cel.MergeCells returns true if the current cell is part of a merged area
                if (cel.MergeCells)
                {
                    //Get the properties of the merge area
                    Range mergedArea = cel.MergeArea;
                    mergedAreas[mergedArea.Address] = mergedArea;
                }
            }

            foreach (Range mergedArea in mergedAreas.Values)
            {
                //Get the properties of the merge area
                string value = mergedArea.FirstCell().Value2;

                if (mode == UnmergeMode.Duplicate)
                {
                    //Unmerge the area and duplicate the value in every cell
                    mergedArea.UnMerge();
                    foreach (Range mergeCel in mergedArea)
                    {
                        mergeCel.Value2 = value;
                    }
                }
                else if (mode == UnmergeMode.SplitByRow)
                {
                    //var rowCount = mergedArea.EntireRow.Count;
                    mergedArea.UnMerge();

                    //The original text, split by line
                    var textLines = value.Split(new string[] { Environment.NewLine, "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    int iRow = 0;
                    foreach (Range row in mergedArea.Rows)
                    {
                        var rowFirstCel = row.FirstCell();
                        row.Clear();
                        row.Merge();
                        if (iRow < textLines.Length)
                        {
                            rowFirstCel.Value2 = textLines[iRow];
                        }
                        iRow++;
                    }

                }

            }
        }

        public enum UnmergeMode
        {
            Duplicate,
            SplitByRow
        }
    }
}
