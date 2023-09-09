using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Forms = System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Linq.Expressions;

namespace Excel4Engineers
{
    /// <summary>
    /// The functions used in the [Name] & [Batch Name] groups of the ribbon
    /// </summary>
    public class NameFunctions
    {

        /*
Rules for valid excel names:
 - Names must begin with a letter, an underscore (_), or a backslash (\)
 - Names can't contain spaces and most punctuation characters.
 - Names can't conflict with cell references – you can't name a range "A1" or "Z100".
 - Single letters are OK for names ("a", "b", "x", etc.), but the letters "r" and "c" are reserved.
 - Names are not case-sensitive – "home", "HOME", and "HoMe" are all the same to Excel.


 */

        public static void NameSelection(bool isWorksheetLevelName = true)
        {
            var selection = RangeHelpers.GetSelection();
            if (selection.Count == 0) return;

            Application xlApp = (Application)ExcelDnaUtil.Application;
            object inputBoxResult = xlApp.InputBox("Enter a name:", isWorksheetLevelName ? "Worksheet Level Name" : "Global Name", Type: 0);

            if (inputBoxResult is bool boolResult && boolResult == false)
            {
                return;
            }

            if (!(inputBoxResult is string nameText))
            {
                throw new Exception("Whyy??");
            }

            if (nameText.StartsWith(@"=""") && nameText.EndsWith(@""""))
            {
                nameText = nameText.Substring(2, nameText.Length - 3);
            }

            bool nameResult = CreateName(isWorksheetLevelName, selection, nameText, out string errorMessage);
            
            if(errorMessage != null)
                Forms.MessageBox.Show($"Could not name {selection.Address} with the name [{nameText}] due to the following error: {errorMessage}");
        }

        private static bool CreateName(bool isWorksheetLevelName, Range range, string nameText, out string errorMessage)
        {
            if (range == null) 
            {
                errorMessage = "Range null.";
                return false;
            }

            if (isWorksheetLevelName)
            {
                var ws = range.Worksheet;
                if(DoesWorksheetNameExist(nameText, ws))
                {
                    errorMessage = $"Worksheet name {nameText} already exists.";
                    return false;
                }

                try
                {
                    ws.Names.Add(Name: nameText, RefersTo: range);
                }
                catch (Exception e)
                {
                    errorMessage = e.Message;
                    return false;
                }
            }
            else
            {
                var wb = range.Worksheet.Workbook();
                if (DoesGlobalNameExist(nameText, wb))
                {
                    errorMessage = $"Name {nameText} already exists.";
                    return false;
                }

                try
                {
                    wb.Names.Add(nameText, RefersToLocal: range);
                }
                catch (Exception e)
                {
                    errorMessage = e.Message;
                    return false;
                }
            }
            errorMessage = null;
            return true;

        }

        public static bool DoesGlobalNameExist(string rangeName, Workbook wb)
        {
            try
            {
                wb.Names.Item(rangeName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool DoesWorksheetNameExist(string rangeName, Worksheet ws)
        {
            try
            {
                ws.Names.Item(rangeName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void DeleteNamesInSelection()
        {
            var selection = RangeHelpers.GetSelection();
            foreach (Name nm in selection.Worksheet.Workbook().Names)
            {
                Range nameRange = nm.RefersToRange;
                if (nameRange != null)
                {
                    if (selection.Worksheet != nameRange.Worksheet) continue;
                    if(selection.Application.Intersect(selection, nameRange) != null)
                    {
                        nm.Delete();
                    }
                }
            }
            foreach (Name nm in selection.Worksheet.Names)
            {
                Range nameRange = nm.RefersToRange;
                if (nameRange != null)
                {
                    if (selection.Worksheet != nameRange.Worksheet) continue;
                    if (selection.Application.Intersect(selection, nameRange) != null)
                    {
                        nm.Delete();
                    }
                }
                
            }
        }

        public static void BatchNameInSelection(int rowColumnOffset, bool isWorksheetLevelName = true)
        {
            var selection = RangeHelpers.GetSelection();
            if (selection == null) return;

            //Check the count of rows and columns
            int rows = selection.Rows.Count;
            int columns = selection.Columns.Count;

            if (rows != 1 && columns != 1) 
            {
                Forms.MessageBox.Show("The selection for batch name must be either [1Row x nCols] or [nRows x 1Col]. The names will be taken from the offset cells [above] or [to the left] respectively.");
                return;
            }

            var firstCell = selection.FirstCell();
            var errorMessages = new HashSet<string>();
            var failedNameCells = new List<string>();            
            //When there is a one row selected, and the offset names are above
            if (rows == 1)
            {
                //Check that the offset cell is within bounds of the worksheet
                if(firstCell.Row - rowColumnOffset <= 0)
                {
                    Forms.MessageBox.Show($"The offset range (containing the proposed names) is not valid; the row offset cannot be applied to the selected name because the row no. would be {firstCell.Row - rowColumnOffset}.");
                    return;
                }

                Range rangeWithName = selection.Offset[-rowColumnOffset, 0];
                //var x = rangeWithName.Address;
                var rangeWithNameCropped = rangeWithName.CropToUsedRange();
                //var y = rangeWithNameCropped.Address;

                foreach (Range celWithName in rangeWithNameCropped)
                {
                    var proposedName = celWithName.Text;
                    if (string.IsNullOrEmpty(proposedName)) continue;

                    var celToBeNamed = celWithName.Offset[rowColumnOffset, 0];
                    var nameResult = CreateName(isWorksheetLevelName, celToBeNamed, proposedName, out string errorMessage);
                    if (!nameResult)
                    {
                        errorMessages.Add(errorMessage);
                        failedNameCells.Add($@"- {celToBeNamed.Address} : ""{proposedName}""");
                    }
                }
            }
            //When there is a one column selected, and the offset names are to the left
            else
            {
                //Check that the offset cell is within bounds of the worksheet
                if (firstCell.Column - rowColumnOffset <= 0)
                {
                    Forms.MessageBox.Show($"The offset range (containing the proposed names) is not valid; the column offset cannot be applied to the selected name because the column no. would be {firstCell.Column - rowColumnOffset}.");
                    return;
                }

                Range rangeWithName = selection.Offset[0, -rowColumnOffset];
                //var x = rangeWithName.Address;
                var rangeWithNameCropped = rangeWithName.CropToUsedRange();
                //var y = rangeWithNameCropped.Address;

                foreach (Range celWithName in rangeWithNameCropped)
                {
                    var proposedName = celWithName.Text;
                    if (string.IsNullOrEmpty(proposedName)) continue;

                    var celToBeNamed = celWithName.Offset[0,rowColumnOffset];
                    var nameResult = CreateName(isWorksheetLevelName, celToBeNamed, proposedName, out string errorMessage);
                    if (!nameResult)
                    {
                        errorMessages.Add(errorMessage);
                        failedNameCells.Add($@"- {celToBeNamed.Address} : ""{proposedName}""");
                    }
                }
            }

            if(failedNameCells.Count > 0)
            {
                Forms.MessageBox.Show($@"Failed to name {failedNameCells.Count} cells:
{string.Join(Environment.NewLine, failedNameCells.ToArray())}

{string.Join(Environment.NewLine+Environment.NewLine, errorMessages.ToArray())}");
            }

            ////Offset the selected range, and crop it to the used range (in case the user selected the entire column)
            //Range nameRange;
            ////When there is a one row selected, and the offset names are above
            //if (rows == 1)
            //{
            //    nameRange = selection.Offset[-rowColumnOffset, 0];
            //}
            ////When there is a one column selected, and the offset names are to the left
            //else
            //{
            //    nameRange = selection.Offset[0,-rowColumnOffset];
            //} 

        }
    }
}
