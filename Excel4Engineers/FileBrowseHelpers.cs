using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel4Engineers
{
    public static class FileBrowseHelpers
    {
        /// <summary>
        /// Opens a FolderBrowser dialog to get a folderpath
        /// </summary>
        public static string GetFolderpath(string initialFolderpath = null)
        {
            string folderpath = initialFolderpath;
            if (string.IsNullOrEmpty(initialFolderpath) || !Directory.Exists(initialFolderpath))
                folderpath = "c:\\";

            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = folderpath;
                fbd.RootFolder = Environment.SpecialFolder.Desktop;

                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    return fbd.SelectedPath;
                }
                return initialFolderpath;
            }
        }
    }
}
