using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    /// <summary>
    /// A helper class for getting data from a Template sheet like for the FileScheduleFunctions
    /// These functions rely on certain named cells within a template sheet.
    /// This helper class is used to retrieve the data within the required ranges/listObjects for a function,
    /// and creates error messages if these ranges are missing.
    /// </summary>
    public class ActiveWorksheetDataLoader
    {
        /// <summary>
        /// Status of getting the active worksheet.
        /// null - not required
        /// false - tried and failed to get
        /// true - available
        /// </summary>
        public bool? ActiveworksheetStatus { get; private set; }

        /// <summary>
        /// The activeworkbook
        /// </summary>
        public Workbook ActiveWorkbook { get; private set; }

        /// <summary>
        /// The active worksheet
        /// </summary>
        public Worksheet ActiveWorksheet { get; private set; }

        /// <summary>
        /// The first listobject in the active sheet. This is only cached if TryGetFirstListobject has been called
        /// </summary>
        public ListObject FirstListobject { get; private set; }

        /// <summary>
        /// Status of getting the active worksheet.
        /// null - not required
        /// false - tried and failed to get
        /// true - available
        /// </summary>
        public bool? ListobjectStatus { get; private set; }

        /// <summary>
        /// Shows true if all required inputs were found.
        /// </summary>
        public bool Success { get; private set; } = true;

        /// <summary>
        /// The list of warnings accumlated during the data-retrieval process
        /// </summary>
        public List<string> Warnings { get; private set; } = new List<string>();

        /// <summary>
        /// Create an ActiveWorksheetDataLoader. This automatically gets the ActiveWorksheet during initialisation.
        /// </summary>
        public ActiveWorksheetDataLoader()
        {
            TryGetActiveWorksheet();

        }

        public bool TryGetActiveWorksheet()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            ActiveWorkbook = xlApp.ActiveWorkbook;
            if (ActiveWorkbook == null)
            {
                ActiveworksheetStatus = false;
                Success = false;
                Warnings.Add("An active workbook is required to run this function.");
                return false;
            }

            ActiveWorksheet = ActiveWorkbook.ActiveSheet as Worksheet;
            ActiveworksheetStatus = ActiveWorksheet != null;
            return ActiveworksheetStatus.Value;
        }

        /// <summary>
        /// Tries to get a listObject from the active worksheet, logging errors where applicable
        /// </summary>
        public bool TryGetFirstListobject(out ListObject listobject)
        {
            if(ActiveworksheetStatus == null)
            {
                TryGetActiveWorksheet();
            }

            if (ActiveworksheetStatus != true) 
            {
                listobject = null;
                Success = false;
                Warnings.Add("An active worksheet is required to run this function.");
                return false;
            }

            var listObjectCount = ActiveWorksheet.ListObjects.Count;
            if(listObjectCount == 0)
            {
                listobject = null;
                ListobjectStatus = Success = false;
                Warnings.Add("This function requires the active worksheet to have a Table range. The first one will be used.");
                return false;
            }

            FirstListobject = listobject = ActiveWorksheet.ListObjects[1];
            if (listobject == null)
            {
                ListobjectStatus = Success = false;
                throw new Exception("The first listobject is null! Investigate why this is!");
                return false;
            }

            if (listObjectCount > 1)
            {
                Warnings.Add($"More than one Excel table in the active worksheet. The first table {listobject.Name} will be used.");
            }

            ListobjectStatus = true;
            return true;
        }

        /// <summary>
        /// Tries to get a list column from the active worksheet, logging errors where applicable
        /// </summary>
        public bool TryGetListobjectColumn(string columnName, out ListColumn listcolumn)
        {
            if (ListobjectStatus == null)
            {
                TryGetFirstListobject(out ListObject lo);
            }

            if (ListobjectStatus != true)
            {
                listcolumn = null;
                return false;
            }

            listcolumn = FirstListobject.GetColumn(columnName);
            if(listcolumn == null)
            {
                Warnings.Add($"Table column {columnName} could not be found.");
                Success = false;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Tries to get a range from the active worksheet, logging errors where applicable
        /// </summary>
        public bool TryGetActiveWorksheetRange(string rangeAddress, out Range range)
        {
            if (ActiveworksheetStatus == null)
            {
                TryGetActiveWorksheet();
            }

            if(!ActiveWorksheet.TryGetRange(rangeAddress, out range))
            {
                Warnings.Add($"Range {rangeAddress} could not be found.");
                Success = false;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates the warning text for any warnings accumalated during initilsation of this ActiveWorksheetDataLoader
        /// </summary>
        public string PrintWarnings(string preamble)
        {
            return preamble + Environment.NewLine + string.Join(Environment.NewLine, Warnings.ToArray());
        }

        /// <summary>
        /// Shows the Messagebox with any warning texts accumalated during initilsation of this ActiveWorksheetDataLoader
        /// </summary>
        public void ShowMessageboxWithWarnings(string preamble)
        {
            System.Windows.Forms.MessageBox.Show(PrintWarnings(preamble));
        }
    }
}
