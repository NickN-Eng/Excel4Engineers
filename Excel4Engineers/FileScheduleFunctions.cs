using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    /// <summary>
    /// The functions used in the [File Schedule] group of the ribbon
    /// </summary>
    public class FileScheduleFunctions
    {
        /// <summary>
        /// Inserts the file template
        /// </summary>
        public static void OpenFileTemplate()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            if (app?.ActiveWorkbook == null)
            {
                app.Workbooks.Add(Type.Missing);

            }

            ExcelResourcesTemplateLoader.OpenExcelTemplateFromResources(Properties.Resources.Template_File, "FileTemplate", app.ActiveWorkbook);
        }

        /// <summary>
        /// Load the files in the Folderpath directory into the FullPaths column of the template
        /// </summary>
        public static void LoadFileData()
        {
            var dataLoader = new ActiveWorksheetDataLoader();

            dataLoader.TryGetActiveWorksheetRange("Folderpath", out Range folderPathRange);
            dataLoader.TryGetActiveWorksheetRange("SearchSubfolders", out Range searchSubfoldershRange);
            dataLoader.TryGetListobjectColumn("FullPath", out ListColumn fullPathColumn);
            dataLoader.TryGetListobjectColumn("FileName", out ListColumn fileNameColumn);
            dataLoader.TryGetListobjectColumn("Ext.", out ListColumn extColumn);
            dataLoader.TryGetListobjectColumn("Date Mod.", out ListColumn dateModColumn);
            
            if (!dataLoader.Success)
            {
                dataLoader.ShowMessageboxWithWarnings("Could load file data.");
            }

            var fullPaths = new List<string>();
            var fileName = new List<string>();
            var extension = new List<string>();

            var folderpath = folderPathRange.Value2;
            if(!Directory.Exists(folderpath))
            {
                System.Windows.Forms.MessageBox.Show("Folderpath is invalid. Could not load file list.");
                return;
            }

            string searchSubfolders = searchSubfoldershRange.Value2;
            var searchSubBool = searchSubfolders.ParseTextToBool();

            var so = searchSubBool.HasValue ? (searchSubBool.HasValue ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly) : SearchOption.TopDirectoryOnly;
            foreach (var path in Directory.GetFiles(folderpath, "*", so))
            {
                fullPaths.Add(path);
                fileName.Add(System.IO.Path.GetFileNameWithoutExtension(path));
                extension.Add(System.IO.Path.GetExtension(path));

                //string dateValue = null;
                //try
                //{
                //    dateValue = System.IO.File.time GetLastWriteTime(path).ToString();
                //}
                //catch (System.Exception) { }
                //date.Add(dateValue);
            }

            fullPathColumn.DataBodyRange.WriteValuesAsColumn(fullPaths);
            fileNameColumn.DataBodyRange.WriteValuesAsColumn(fileName);
            extColumn.DataBodyRange.WriteValuesAsColumn(extension);
        }

        /// <summary>
        /// Run the ScanPdfs algorithm and write the results to the results columns
        /// </summary>
        public static void LoadPdfData()
        {
            //Load data from the dataloader
            var dataLoader = new ActiveWorksheetDataLoader();
            dataLoader.TryGetActiveWorksheetRange("PdfTemplate", out Range templatePathRange);
            dataLoader.TryGetFirstListobject(out ListObject firstListobject);
            dataLoader.TryGetListobjectColumn("FullPath", out ListColumn fullPathColumn);

            //Check there pdf template is a valid file
            var templatePath = templatePathRange.Value2;
            if (!File.Exists(templatePath))
            {
                System.Windows.Forms.MessageBox.Show("Pdf template is invalid. Could not load file data.");
                return;
            }

            //Set up the pdf reader
            PdfReader pdfReader = new PdfReader();
            bool templateSuccess = pdfReader.SetTemplate(templatePath);
            if (templateSuccess)
            {
                System.Windows.Forms.MessageBox.Show("Pdf template is invalid. Could not load file data.");
                return;
            }

            //Show any pdf template messages if required
            if (pdfReader.TemplateMessages.Count > 0)
            {
                pdfReader.ShowMessageboxTemplateMessages($"Read pdf template has {pdfReader.TemplateMessages.Count} notifications:");
            }

            //Initilaise arrays to contain the search results
            var numberOfSearchRegions = pdfReader.SearchBoxes.Count;
            var filepaths = fullPathColumn.DataBodyRange.GetValuesAsString();
            List<string[]> searchColumnResults = new List<string[]>();
            for (int i = 0; i < numberOfSearchRegions; i++)
            {
                searchColumnResults.Add(Enumerable.Repeat("", filepaths.Length).ToArray());
            }

            //Scan each pdf in the filepaths list & write the results in the relevant column
            for (int i = 0; i < filepaths.Length; i++)
            {
                var filepath = filepaths[i];

                if (Path.GetExtension(filepath) != ".pdf") continue;

                pdfReader.ScanPdf(filepath, out List<string> foundText);

                for (int j = 0; j < numberOfSearchRegions; j++)
                {
                    searchColumnResults[j][i] = foundText[j];
                }
            }

            //Insert the columns with the scan results after the last column of the File Schedule table
            for (int i = 0; i < numberOfSearchRegions; i++)
            {
                firstListobject.InsertColumn(pdfReader.SearchBoxes[i].Name, searchColumnResults[i]);
            }

            //Show any relevant warnings
            if(pdfReader.Warnings.Count > 0)
            {
                pdfReader.ShowMessageboxWithWarnings($"Pdf scan complete with {pdfReader.Warnings.Count} warnings:");
            }

        }

        /// <summary>
        /// Copies or moves files in the Fullpath column according to the directory in the CopyMovePath folder
        /// </summary>
        public static void CopyMoveFileData(bool copy = true)
        {
            var dataLoader = new ActiveWorksheetDataLoader();

            dataLoader.TryGetListobjectColumn("FullPath", out ListColumn fullPathColumn);
            dataLoader.TryGetListobjectColumn("CopyMovePath", out ListColumn copyMoveColumn);

            if (!dataLoader.Success)
            {
                dataLoader.ShowMessageboxWithWarnings("Could load file data.");
            }

            var fullPaths = fullPathColumn.DataBodyRange.GetValuesAsString();
            var copyMovePaths = copyMoveColumn.DataBodyRange.GetValuesAsString();

            var messages = new List<string>();
            var sucessCount = 0;

            for (int i = 0; i < fullPaths.Length; i++)
            {
                var fullPath = fullPaths[i];
                var copyMovePath = copyMovePaths[i];

                //Check that the file to be copied exists
                if (!File.Exists(fullPath))
                {
                    messages.Add($"File at: {fullPath} does not exist. Could not {(copy ? "copy" : "move")}.");
                    continue;
                }

                if (string.IsNullOrEmpty(copyMovePath))
                {
                    continue;
                }

                //Check whether the destination path is a file or directory
                //Get the directory and create it/ensure it exists
                //If the copyMovePath is just a directory, use the current filename
                string directory = copyMovePath;
                string combinedCopyMovePath = copyMovePath;
                FileAttributes attr = File.GetAttributes(copyMovePath);
                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                    combinedCopyMovePath = Path.Combine(copyMovePath, Path.GetFileName(fullPath));
                else
                    directory = Path.GetDirectoryName(copyMovePath);

                Directory.CreateDirectory(directory);

                if (copy)
                {
                    try
                    {
                        File.Copy(fullPath, combinedCopyMovePath);
                        sucessCount++;
                    }
                    catch (Exception e)
                    {
                        messages.Add($"File at: {fullPath} could not be copied due to {e.Message}.");
                    }
                }
                else
                {
                    try
                    {
                        File.Move(fullPath, combinedCopyMovePath);
                        sucessCount++;
                    }
                    catch (Exception e)
                    {
                        messages.Add($"File at: {fullPath} could not be moved due to {e.Message}.");
                    }
                }
            }

            System.Windows.Forms.MessageBox.Show($"{(copy ? "Copy" : "Move")} operation complete with {sucessCount}/{fullPaths.Length} sucesses.{Environment.NewLine}{string.Join(Environment.NewLine, messages.ToArray())}");

        }
    }
}
