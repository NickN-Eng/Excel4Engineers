using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    public static class ListObjectHelpers
    {
        /// <summary>
        /// For this table, gets the range by the column name
        /// Test for single value
        /// </summary>
        public static Range GetColumnRange(this ListObject table, string columnName)
        {
            ListColumns cols = table.ListColumns;
            foreach (ListColumn col in table.ListColumns)
            {
                if (col.Name != columnName) continue;
                return col.DataBodyRange;
            }
            return null;
        }

        /// <summary>
        /// For this table, gets the ListColumn by the column name
        /// </summary>
        public static ListColumn GetColumn(this ListObject table, string columnName)
        {
            ListColumns cols = table.ListColumns;
            foreach (ListColumn col in table.ListColumns)
            {
                if (col.Name != columnName) continue;
                return col; 
            }
            return null;

        }

        /// <summary>
        /// Inserts a column into the table after the last column
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <param name="columnName"></param>
        /// <param name="columnValues"></param>
        public static void InsertColumn<T>(this ListObject table, string columnName, IList<T> columnValues) 
        {
            ListColumns cols = table.ListColumns;

            //var headerRangeAddDebug = table.HeaderRowRange.Address;

            var lastCell = table.HeaderRowRange.BottomRightCell();
            //var lastCellAddDebug = lastCell.Address;

            lastCell.Offset[0,1].EntireColumn.Insert(1);
            
            var headerCell = table.HeaderRowRange.BottomRightCell().Offset[0,1];
            headerCell.Value2 = columnName;

            headerCell.Offset[1, 0].WriteValuesAsColumn(columnValues);
        }

        /// <summary>
        /// Adds a table
        /// </summary>
        public static void AddTable(Range firstCell, List<string> Headers, List<List<string>> ValuesByColumn)
        {
            Worksheet ws = firstCell.Worksheet;
            
            if(Headers.Count != ValuesByColumn.Count)
            {
                //Header and ValuesByColumns list must be same length
                return;
            }

            int rowCount = ValuesByColumn.First().Count;
            foreach (var col in ValuesByColumn)
            {
                if(col.Count != rowCount)
                {
                    //Unequal row count error message
                    return;
                }
            }

            //Write the headers
            firstCell.WriteValuesToRow(Headers);
            //var x = firstCell.Address;

            //Write the values
            var firstCellOfValues = firstCell.Offset[1, 0];
            //var y = firstCellOfValues.Address;

            foreach (var col in ValuesByColumn)
            {
                firstCellOfValues.WriteValuesAsColumn(col);
                firstCellOfValues = firstCellOfValues.Offset[0, 1];
                //var z = firstCellOfValues.Address;

            }

            //var lastCell = firstCell.Offset[rowCount, Headers.Count - 1];
            var tableRange = firstCell.SetRangeByRowColNo(rowCount, Headers.Count - 1);
            //var a = tableRange.Address;

            ws.ListObjects.AddEx(SourceType: XlListObjectSourceType.xlSrcRange, Source: tableRange, XlListObjectHasHeaders: XlYesNoGuess.xlYes);
        }
    }

}
