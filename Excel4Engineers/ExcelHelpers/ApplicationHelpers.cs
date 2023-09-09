//using ExcelBrowser.Interop;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Excel4Engineers
{
    /// <summary>
    /// Helper methods & extension methods for working with Excel application
    /// </summary>
    public static class ApplicationHelpers
    {
        public static IEnumerable<Workbook> GetWorkbooks(this Application app)
        {
            foreach (Workbook wb in app.Workbooks)
            {
                yield return wb;
            }
        }
    }
}
