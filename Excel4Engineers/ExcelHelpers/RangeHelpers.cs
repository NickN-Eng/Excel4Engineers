using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    /// <summary>
    /// Helper methods & extension methods for working with Excel ranges
    /// </summary>
    public static class RangeHelpers
    {

        /// <summary>
        /// Writes an array to excel based on the row/column position of the top corner cell
        /// Ignores the rest of the range. Cells outside the range can be overwritten based on the size of the array.
        /// Returns the final range.
        /// </summary>
        public static Range WriteGridValues<T>(this Range startCell, T[,] values)
        {
            var worksheet = startCell.Worksheet;
            int row = startCell.Row;
            int col = startCell.Column;
            var cell1 = (Range)worksheet.Cells[row, col];
            var cell2 = (Range)worksheet.Cells[row + values.GetLength(0) - 1, col + values.GetLength(1) - 1];
            Range rng = worksheet.get_Range(cell1, cell2);
            rng.Value = values;
            return rng;
        }

        /// <summary>
        /// Writes a list to excel as a column of values based on the row/column position of the top corner cell
        /// Ignores the rest of the range. Cells outside the range can be overwritten based on the size of the array.
        /// </summary>
        public static Range WriteValuesAsColumn<T>(this Range startCell, IList<T> values)
        {
            var worksheet = startCell.Worksheet;
            int row = startCell.Row;
            int col = startCell.Column;
            var valueArray = new T[values.Count, 1];
            for (int i = 0; i < values.Count; i++)
            {
                valueArray[i, 0] = values[i];
            }
            return WriteGridValues(startCell, valueArray);
        }

        /// <summary>
        /// Writes a list to excel as a row of values based on the row/column position of the top corner cell
        /// Ignores the rest of the range. Cells outside the range can be overwritten based on the size of the array.
        /// </summary>
        public static void WriteValuesToRow<T>(this Range range, IEnumerable<T> values) => WriteValuesToRow(range, values.ToArray());

        /// <summary>
        /// Writes a list to excel as a row of values based on the row/column position of the top corner cell
        /// Ignores the rest of the range. Cells outside the range can be overwritten based on the size of the array.
        /// </summary>
        public static Range WriteValuesToRow<T>(this Range startCell, T[] values)
        {
            var worksheet = startCell.Worksheet;
            int row = startCell.Row;
            int col = startCell.Column;
            var cell1 = (Range)worksheet.Cells[row, col];
            var cell2 = (Range)worksheet.Cells[row, col + values.Length - 1];
            Range rng = worksheet.get_Range(cell1, cell2);
            rng.Value = values;
            return rng;
        }

        /// <summary>
        /// From the first cell of the  startCel range, create a range with the [rows x columns]
        /// </summary>
        /// <param name="startCel"></param>
        /// <param name="rows">Number of rows</param>
        /// <param name="columns">Number of columns</param>
        /// <returns></returns>
        public static Range SetRangeByRowColNo(this Range startCel, int rows, int columns)
        {
            var worksheet = startCel.Worksheet;
            var cell1 = startCel.Cells[1,1];
            var cell2 = startCel.Offset[rows, columns];
            return worksheet.Range[cell1, cell2];
        }

        /// <summary>
        /// Get the first cell of this range
        /// </summary>
        public static Range FirstCell(this Range range)
        {
            return range.Cells[1, 1];
        }

        /// <summary>
        /// Gets the bottom right cell of the range, assuming the range is a single area, rectangular range.
        /// </summary>
        public static Range BottomRightCell(this Range range)
        {
            var firstCol = range.Column;
            var firstRow = range.Row;
            var noColumns = range.EntireColumn.Count;
            var noRows = range.EntireRow.Count;

            var lastCell = range.Worksheet.Cells[firstRow + noRows - 1, firstCol + noColumns - 1];

            return lastCell;
        }

        /// <summary>
        /// Gets Value2 from the range as and array of objects
        /// </summary>
        public static object[] GetValues(this Range range)
        {
            var val = range.Value2;

            if (val is Array array)
                return array.Cast<object>().ToArray();
            else
                return new object[] { val };
        }


        /// <summary>
        /// Gets Value2 from the range as and array of strings
        /// </summary>
        public static string[] GetValuesAsString(this Range range)
        {
            var val = range.Value2;

            if(val is Array array)
            {
                string[] strArray = new string[array.Length];
                int i = 0;
                foreach (var item in array)
                {
                    if (item == null) 
                        strArray[i] = "";
                    else
                        strArray[i] = item.ToString();
                    i++;
                }
                return strArray;
            }
            else
            {
                if (val == null)
                    return new string[] { "" };
                else
                    return new string[] { val.ToString() };
            }
        }

        /// <summary>
        /// Gets Value2 from the range as and array of doubles
        /// Replaces strings, boolean and error types with the errorValue.
        /// </summary>
        public static double[] GetValuesAsDoubles(this Range range, double errorValue = double.NaN)
        {
            var val = range.Value2;

            if (val is Array array)
            {
                //Is this better or worse??   
                double[] dArray = new double[array.Length];
                int i = 0;
                foreach (var item in array)
                {
                    if (item is double d)
                        dArray[i] = d;
                    else
                        dArray[i] = errorValue;
                    i++;
                }
                return dArray;
            }
            else
            {
                if (val is double d)
                    return new double[] { d };
                else
                    return new double[] { errorValue };
            }
        }

        /// <summary>
        /// Returns true of the range is empty by the Range.Value
        /// </summary>
        public static bool IsRangeEmpty(this Range range)
        {
            var values = range.Value;
            if (values == null) return true;

            if (values is object[,] array)
            {
                foreach (var cell in array)
                {
                    if (cell != null) return false;
                }
                return true;
            }

            return (object)values == null;

        }


        /// <summary>
        /// Tries to get the range from a worksheet, by an address
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool TryGetRange(this Worksheet worksheet, string address, out Range range)
        {
            try
            {
                range = worksheet.Range[address];
                return range != null;
            }
            catch (Exception)
            {
                range = null;
                return false;
            }
        }

        /// <summary>
        /// The number of cells in the selection before used range will be be used to crop the selection
        /// </summary>
        public const int ThresholdForUsedRangeLimit = 2000;

        /// <summary>
        /// Gets the user selection
        /// Limit to used range is an option feature which crops the selection to the used range in the worksheet.
        /// This is useful if the algorithm will iterate over every cell of the selection because the user may select
        /// an entire column or sheet, which would otherwise cause the algorithm to loop over 10000's of cells.
        /// The selection will not be cropped unless the selected cell count exceeds the ThresholdForUsedRangeLimit=2000.
        /// </summary>
        /// <param name="limitToUsedRange"></param>
        /// <returns></returns>
        public static Range GetSelection(bool limitToUsedRange = false)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Range selection = xlApp.Selection;

            if (selection == null) return null;
            if (selection.Count == 0) return selection;

            if (limitToUsedRange && selection.Count > ThresholdForUsedRangeLimit)
            {
                var usedRange = selection.Worksheet.UsedRange;
                return xlApp.Intersect(selection, usedRange);
            }
            else
            {
                return selection;
            }
        }

        /// <summary>
        /// Crops a range to the used range within the the worksheet.
        /// This may be needed if the user has selected the entire column, and we do not want to iterate through the entire column which has thousands of cells.
        /// </summary>
        public static Range CropToUsedRange(this Range rng)
        {
            if (rng == null) return null;
            if (rng.Count == 0) return rng;

            var usedRange = rng.Worksheet.UsedRange;
            return rng.Application.Intersect(rng, usedRange);
        }
    }
}
