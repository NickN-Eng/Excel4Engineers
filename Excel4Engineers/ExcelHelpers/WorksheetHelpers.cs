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
    /// Helper methods & extension methods for working with Excel workbooks
    /// </summary>
    public static class WorksheetHelpers
    {
        /// <summary>
        /// Gets the workbook from a worksheet
        /// </summary>
        public static Workbook Workbook(this Worksheet worksheet)
        {
            return worksheet.Parent;
        }

        /// <summary>
        /// Write a value to a cell by the row/column index (indexes start from 1)
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        /// <param name="row">Index, starting at 1 for row 1</param>
        /// <param name="col">Index starting at 1 for column A</param>
        /// <param name="value">Single value to be written</param>
        public static void WriteValueToCell(this Worksheet worksheet, int row, int col, object value)
        {
            var rng = (Range)worksheet.Cells[row, col];
            rng.Value = value;
        }

        /// <summary>
        /// Write a list of values to as a vertical column by the row/column index of a start cell(indexes start from 1)
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        /// <param name="row">Index, starting at 1 for row 1</param>
        /// <param name="col">Index starting at 1 for column A</param>
        /// <param name="values">Single value to be written</param>
        public static void WriteValuesToColumn(this Worksheet worksheet, int row, int col, IList<object> values)
        {
            var valueArray = new object[values.Count, 1];
            for (int i = 0; i < values.Count; i++)
            {
                valueArray[i, 0] = values[i];
            }
            WriteGridValues(worksheet, row, col, valueArray);
        }

        /// <summary>
        /// Write a list of values to as a horizontal row by the row/column index of a start cell(indexes start from 1)
        /// </summary>
        /// <param name="row">Index, starting at 1 for row 1</param>
        /// <param name="col">Index starting at 1 for column A</param>
        /// <param name="values">Single value to be written</param>
        /// <param name="worksheet">Optional worksheet</param>
        public static void WriteValuesToRow(Worksheet worksheet, int row, int col, IEnumerable<object> values) => WriteValuesToRow(worksheet, row, col, values.ToArray());


        /// <summary>
        /// Write a list of values to as a horizontal row by the row/column index of a start cell(indexes start from 1)
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        /// <param name="row">Index, starting at 1 for row 1</param>
        /// <param name="col">Index starting at 1 for column A</param>
        /// <param name="values">Single value to be written</param>
        public static void WriteValuesToRow(Worksheet worksheet, int row, int col, object[] values)
        {
            var cell1 = (Range)worksheet.Cells[row, col];
            var cell2 = (Range)worksheet.Cells[row, col + values.Length - 1];
            Range rng = worksheet.get_Range(cell1, cell2);
            rng.Value = values;
        }


        /// <summary>
        /// Writes an array to excel based on the row/column position of the top corner cell
        /// </summary>
        public static void WriteGridValues(this Worksheet worksheet, int row, int col, object[,] values)
        {
            var cell1 = (Range)worksheet.Cells[row, col];
            var cell2 = (Range)worksheet.Cells[row + values.GetLength(0) - 1, col + values.GetLength(1) - 1];
            Range rng = worksheet.get_Range(cell1, cell2);
            rng.Value = values;
        }

        /// <summary>
        /// Gets a range, using a start cell defined by a row/col index (indexes start from 1)
        /// and the width and height of the row
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        /// <param name="row">Index, starting at 1 for row 1</param>
        /// <param name="col">Index starting at 1 for column A</param>
        /// <param name="values">Single value to be written</param>
        public static Range GetRange(this Worksheet worksheet, int row, int col, int height = 1, int width = 1)
        {
            var cell1 = (Range)worksheet.Cells[row, col];
            var cell2 = (Range)worksheet.Cells[row + height - 1, col + width - 1];
            return worksheet.get_Range(cell1, cell2);
        }
    }
}
