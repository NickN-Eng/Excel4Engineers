using Excel4Engineers.Properties;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel4Engineers.WPFControls;

namespace Excel4Engineers
{
    /// <summary>
    /// Provides the Excel4Engineers ribbon within excel
    /// Generates the xml string which is used by Excel to show the ribbon and provides method bindings for buttons etc...
    /// Methods and properties are shown in this order within the ribbon, and named according to the attributes attatched.
    /// </summary>
    [ComVisible(true)]
    public class RibbonController : RibbonControllerBase
    {
        public override string TabName => "E4Engineers";

        #region Formulae text

        [Button("Formulae Info", "Formulae Text", iconImage:nameof(Resources.Info))]
        public void Formulae_ShowHelpImage(IRibbonControl control)
        {
            ImageDisplay.ShowImageDialog("Formulae text help info", Resources.SuperSubscriptSymbolApply, 0.7);
        }


        [Checkbox("SubSupercript", "Formulae Text", tooltipHeader: "Enable Math symbol replacement", tooltipDescription: SubSuperDescription)]
        public DynamicBool SubSupercript { get; set; } = new DynamicBool(true);


        [Checkbox("Math Symbols", "Formulae Text", tooltipHeader: "Enable Math symbol replacement", tooltipDescription: MathSymbolDescription)]
        public DynamicBool MathSymbols { get; set; } = new DynamicBool(true);


        public const string SubSuperDescription = @"Format subscript (with ""_"") and superscript (with ""^"").";
        public const string MathSymbolDescription = @"Replaces keywords with thier symbolic counterpart:
\Alpha => Α;  
\Beta => Β;  
\Gamma => Γ;  
\Delta => Δ;  
\Epsilon => Ε;  
\Zeta => Ζ;  
\Eta => Η;  
\Theta => Θ;  
\Iota => Ι;  
\Kappa => Κ;  
\Lambda => Λ;  
\Mu => Μ;  
\Nu => Ν;  
\Xi => Ξ;  
\Omicron => Ο;  
\Pi => Π;  
\Rho => Ρ;  
\Sigma => Σ;  
\Tau => Τ;  
\Upsilon => Υ;  
\Phi => Φ;  
\Chi => Χ;  
\Psi => Ψ;  
\Omega => Ω;  
\alpha => α;  
\beta => β;  
\gamma => γ;  
\delta => δ;  
\epsilon => ε;  
\zeta => ζ;  
\eta => η;  
\theta => θ;  
\iota => ι;  
\kappa => κ;  
\lambda => λ;  
\mu => μ;  
\nu => ν;  
\xi => ξ;  
\omicron => ο;  
\pi => π;  
\rho => ρ;  
\sigma => σ;  
\tau => τ;  
\upsilon => υ;  
\phi => φ;  
\chi => χ;  
\psi => ψ;  
\omega => ω;  

";



        [Button("Revert", "Formulae Text", true, "Revert formulae text changes", "Revert changes made by the [Apply] button.", iconImage:nameof(Resources.RevertSubSuperSymbols))]
        public void Formulae_ApplySubscriptSuperscript(IRibbonControl control)
        {
            //var txt = Formulae_ComboboxOption_Selected.Text;
            var subsuper = SubSupercript.Value;
            var symbols = MathSymbols.Value;

            MathFormatFunctions.RevertSubscriptSuperscript(RangeHelpers.GetSelection(), subsuper, symbols);
        }


        [Button("Apply", "Formulae Text", isLarge: true, tooltipHeader: "Apply formulae text changes", tooltipDescription: "Apply formulae text changes to the selected cells, using Sub Super and/or Math Symbol modes if ticked.", iconImage: nameof(Resources.ApplySubSuperSymbols))]
        public void Formulae_Apply(IRibbonControl control)
        {
            var subsuper = SubSupercript.Value;
            var symbols = MathSymbols.Value;

            MathFormatFunctions.FormatSubscriptSuperscript(RangeHelpers.GetSelection(), subsuper, symbols);
        }

        #endregion


        #region Names

        [Button("Worksheet name", "Names", isLarge: true, tooltipHeader: "Create a worksheet-level name", tooltipDescription: "For the selected cells, create one name with a worksheet level scope, using a dialog box.", iconImage: nameof(Resources.WorksheetNames))]
        public void Names_NameWorksheetLevel(IRibbonControl control)
        {
            NameFunctions.NameSelection();
        }

        [Button("Delete names", "Names", true, tooltipHeader:"Delete names", tooltipDescription:"Delete names in the selected cells.", iconImage: nameof(Resources.DeleteNames))]
        public void Names_DeleteNames(IRibbonControl control)
        {
            NameFunctions.DeleteNamesInSelection();
        }

        [Button("Batch Name Info", "Batch Names", iconImage: nameof(Resources.Info))]
        public void Names_ShowHelpImage(IRibbonControl control)
        {
            ImageDisplay.ShowImageDialog("Batch name help info", Resources.BatchNameHelp, 0.7);
        }

        [Combobox("  Type", "Batch Names", new string[] { "Worksheet", "Workbook" })]
        public DynamicText Names_ComboboxType { get; set; } = new DynamicText("Worksheet");


        [Combobox("  Offset", "Batch Names", new string[] { "1", "2", "3", "4", "5", "6" })]
        public DynamicText Names_ComboboxOffset { get; set; } = new DynamicText("1");


        [Button("Batch  Name", "Batch Names", true, tooltipHeader:"Name a batch of cells", tooltipDescription: "For the selected cells, give each cell a name with a worksheet level scope, using the offset as shown in the [Help info] button.", iconImage:nameof(Resources.BatchName))]
        public void Names_BatchWorkbookName(IRibbonControl control)
        {
            int offset = int.Parse(Names_ComboboxOffset.Text);
            bool isWorksheetLevel = Names_ComboboxType.Text == "Worksheet";

            NameFunctions.BatchNameInSelection(offset, isWorksheetLevel);
        }


        #endregion


        #region File schedule group


        [Button("File schedule Info", "File Schedule", iconImage: nameof(Resources.Info))]
        public void File_ShowHelpImage(IRibbonControl control)
        {
            MultiImageDisplay.ShowImageDialog("File schedule help info", new System.Drawing.Image[] { Resources.FileSchedule_1, Resources.FileSchedule_2, Resources.FileSchedule_3, Resources.FileSchedule_4 }, 0.8);
        }

        [Button("Load Template", "File Schedule", tooltipHeader: "Load template for file schedule processes", tooltipDescription: "Inserts the template worksheet for the File Schedule processes into the current workbook. Click the info button for further details.", iconImage: nameof(Resources.Numbers_1))]
        public void File_LoadTemplate(IRibbonControl control)
        {
            FileScheduleFunctions.OpenFileTemplate();
        }

        [Button("Load File Data", "File Schedule", tooltipHeader: "Load file data onto template", tooltipDescription: "Scans the files at the [Folderpath] cell of the template. Click the info button for further details.", iconImage: nameof(Resources.Numbers_2))]
        public void File_LoadFileList(IRibbonControl control)
        {
            FileScheduleFunctions.LoadFileData();
        }

        [Button("Scan Pdfs", "File Schedule", tooltipHeader: "Scan pdf title block text", tooltipDescription: "Using the [PdfTemplate] specified in the template; takes the pdf annotations and uses these to read the title block text for each pdf in the [FullPath] cells. Click the info button for further details.", iconImage: nameof(Resources.Numbers_3))]
        public void File_ScanPdfs(IRibbonControl control)
        {
            FileScheduleFunctions.LoadPdfData();
        }

        [Button("Copy", "File Schedule", tooltipHeader: "Copy files", tooltipDescription: "Takes the file at the [FullPath] cells of the template and copies them into the location specified by [CopyMovePath]. Click the info button for further details.", iconImage: nameof(Resources.Numbers_3))]
        public void File_Copy(IRibbonControl control)
        {
            FileScheduleFunctions.CopyMoveFileData(true);
        }

        [Button("Move", "File Schedule", tooltipHeader: "Rename files", tooltipDescription: "Takes the file at the [FullPath] cells of the template and renames them them according to the [RenameName] cells. Click the info button for further details.", iconImage: nameof(Resources.Numbers_3))]
        public void File_Move(IRibbonControl control)
        {
            FileScheduleFunctions.CopyMoveFileData(false);
        }

        #endregion

        #region Merge unmerge


        [Button("Unmerge and duplicate", "Merge and Unmerge", true, tooltipHeader: "Unmerge cells and duplicate", tooltipDescription: "Unmerges any merged cells in the selection, and replicates the text in the original merged cell accross all unmerged cells.", iconImage: nameof(Resources.MergeAndDuplicate))]
        public void MergeUnmerge_UnmergeAndDuplicate(IRibbonControl control)
        {
            MergeUnmergeFunctions.Unmerge(RangeHelpers.GetSelection(), MergeUnmergeFunctions.UnmergeMode.Duplicate);
        }


        [Button("Unmerge Split lines", "Merge and Unmerge", false, tooltipHeader: "Unmerge cells and split by line", tooltipDescription: "Unmerges any merged cells in the selection. If the merged cell contained any text with more than one line, the first cell row gets the text in the first line, second cell row gets the text from the second line, etc...")]
        public void MergeUnmerge_UnmergeAndSplit(IRibbonControl control)
        {
            MergeUnmergeFunctions.Unmerge(RangeHelpers.GetSelection(), MergeUnmergeFunctions.UnmergeMode.SplitByRow);
        }

        [Button("Merge without delete", "Merge and Unmerge", false, tooltipHeader: "Merge cells and without deleting text", tooltipDescription: "Merges all cells in the selection. Previous text in individual cells is joined rather than deleted.")]
        public void MergeUnmerge_MergeWithoutDelete(IRibbonControl control)
        {
            MergeUnmergeFunctions.MergeWithoutDelete(RangeHelpers.GetSelection());
        }


        [Button("Merge Join Across", "Merge and Unmerge", false, tooltipHeader: "Merge and join across rows", tooltipDescription: "Merges cells in the selection row by row, joining text in the same row.")]
        public void MergeUnmerge_MergeAcrossJoin(IRibbonControl control)
        {
            MergeUnmergeFunctions.MergeAcrossJoin(RangeHelpers.GetSelection());
        }

        #endregion


        [DynamicButtonLabel("Help", isLarge: true, tooltipHeader: "Text", tooltipDescription: "Description", iconImage: nameof(Resources.Info))]
        public DynamicText Names_OffsetLabel3 { get; set; } = new DynamicText("Need Help? Hover on buttons");


    }


}
