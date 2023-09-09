using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;

namespace Excel4Engineers
{
    public class ExcelResourcesTemplateLoader
    {
        /// <summary>
        /// Loads an excel worksheet sheet from the resources file, and copies it into an open workbook
        /// </summary>
        public static void OpenExcelTemplateFromResources(byte[] resourcesFile, string sheetName, Workbook targetWorkbook)
        {
            //Create a temp excel file from the resourcesFil byte stream
            string tempPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(tempPath, resourcesFile);

            //Opens the temp excel file
            Application app = (Application)ExcelDnaUtil.Application;
            Workbook tempWb = app.Workbooks.Open(tempPath);
            Worksheet tempSheet = tempWb.Worksheets[sheetName];

            //Copies into the target workbook
            Sheets ws = targetWorkbook.Worksheets;
            tempSheet.Copy(ws[1]);

            //Close the temp workbook
            tempWb.Close();
        }


    }


}
