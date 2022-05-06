using System;
using System.Windows.Forms;
using System.IO;
using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace SpreadsheetExamples {
    public partial class Form1 : Form {

        #region #CreateWorkbook
        // Create a new Workbook object.
        Workbook workbook = new Workbook();
        #endregion #CreateWorkbook

        public Form1() {
            InitializeComponent();
            InitTreeListControl();
            workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic;
        }

        void InitTreeListControl() {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }
        void InitData(GroupsOfSpreadsheetExamples examples) {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Worksheet"));
            examples.Add(new SpreadsheetNode("Rows and Columns"));
            examples.Add(new SpreadsheetNode("Cells"));
            examples.Add(new SpreadsheetNode("Formulas"));
            examples.Add(new SpreadsheetNode("Formatting"));
            examples.Add(new SpreadsheetNode("Import"));
            examples.Add(new SpreadsheetNode("Export"));
            examples.Add(new SpreadsheetNode("Printing"));
            examples.Add(new SpreadsheetNode("Document Properties"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Worksheet" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Active Worksheet", WorksheetActions.AssignActiveWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("New Worksheet", WorksheetActions.AddWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Delete a Worksheet", WorksheetActions.RemoveWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Rename a Worksheet", WorksheetActions.RenameWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Copy a Worksheet within a Workbook", WorksheetActions.CopyWorksheetWithinWorkbookAction));
            examples[0].Groups.Add(new SpreadsheetExample("Copy a Worksheet between Workbooks", WorksheetActions.CopyWorksheetBetweenWorkbooksAction));
            examples[0].Groups.Add(new SpreadsheetExample("Move a Worksheet", WorksheetActions.MoveWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Show/Hide a Worksheet", WorksheetActions.ShowHideWorksheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Show/Hide Gridlines", WorksheetActions.ShowHideGridlinesAction));
            examples[0].Groups.Add(new SpreadsheetExample("Show/Hide Row and Column Headings", WorksheetActions.ShowHideHeadingsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Page Setup (View Type, Page Orientation, Page Margins, Paper Size)", WorksheetActions.PageSetupAction));
            examples[0].Groups.Add(new SpreadsheetExample("Zoom a Worksheet", WorksheetActions.ZoomWorksheetAction));

            // Add nodes to the "Rows and Columns" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("New Row/Column", RowAndColumnActions.InsertRowsColumnsAction));
            examples[1].Groups.Add(new SpreadsheetExample("Delete a Row/Column", RowAndColumnActions.DeleteRowsColumnsAction));
            examples[1].Groups.Add(new SpreadsheetExample("Copy a Row/Column", RowAndColumnActions.CopyRowsColumnsAction));
            examples[1].Groups.Add(new SpreadsheetExample("Show or Hide a Row/Column", RowAndColumnActions.ShowHideRowsColumnsAction));
            examples[1].Groups.Add(new SpreadsheetExample("Row Height and Column Width", RowAndColumnActions.SpecifyRowsHeightColumnsWidthAction));
            examples[1].Groups.Add(new SpreadsheetExample("Group Rows/Columns", RowAndColumnActions.GroupRowsColumnsAction));

            // Add nodes to the "Cells" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Cell Value", CellActions.ChangeCellValueAction));
            examples[2].Groups.Add(new SpreadsheetExample("Cell Value From Text", CellActions.SetValueFromTextAction));
            examples[2].Groups.Add(new SpreadsheetExample("Named Ranges", CellActions.CreateNamedRangeAction));
            examples[2].Groups.Add(new SpreadsheetExample("Add a Hyperlink to a Cell", CellActions.AddHyperlinkAction));
            examples[2].Groups.Add(new SpreadsheetExample("Copy Data Only, Style Only, or Data with Style", CellActions.CopyCellDataAndStyleAction));
            examples[2].Groups.Add(new SpreadsheetExample("Merge/Split Cells", CellActions.MergeAndSplitCellsAction));
            examples[2].Groups.Add(new SpreadsheetExample("Clear Cells", CellActions.ClearCellsAction));

            // Add nodes to the "Formulas" group of examples. 
            examples[3].Groups.Add(new SpreadsheetExample("Constants and Calculation Operators in Formulas", FormulaActions.UseConstantsAndCalculationOperatorsInFormulasAction));
            examples[3].Groups.Add(new SpreadsheetExample("R1C1 References in Formulas", FormulaActions.R1C1ReferencesInFormulassAction));
            examples[3].Groups.Add(new SpreadsheetExample("Names in Formulas", FormulaActions.UseNamesInFormulasAction));
            examples[3].Groups.Add(new SpreadsheetExample("Create Named Formulas", FormulaActions.CreateNamedFormulasAction));
            examples[3].Groups.Add(new SpreadsheetExample("Functions in Formulas", FormulaActions.UseFunctionsInFormulasAction));
            examples[3].Groups.Add(new SpreadsheetExample("Shared and Array Formulas", FormulaActions.CreateSharedAndArrayFormulasAction));

            // Add nodes to the "Formatting" group of examples.
            examples[4].Groups.Add(new SpreadsheetExample("Apply a Style", FormattingActions.CreateModifyApplyStyleAction));
            examples[4].Groups.Add(new SpreadsheetExample("Individual Cell Formatting", FormattingActions.FormatCellAction));
            examples[4].Groups.Add(new SpreadsheetExample("Date Formats", FormattingActions.SetDateFormatsAction));
            examples[4].Groups.Add(new SpreadsheetExample("Number Formats", FormattingActions.SetNumberFormatsAction));
            examples[4].Groups.Add(new SpreadsheetExample("Cell Colors and Background", FormattingActions.ChangeCellColorsAction));
            examples[4].Groups.Add(new SpreadsheetExample("Cell Gradient Fill", FormattingActions.ChangeCellGradientFillAction));
            examples[4].Groups.Add(new SpreadsheetExample("Font Settings", FormattingActions.SpecifyCellFontAction));
            examples[4].Groups.Add(new SpreadsheetExample("Cell Alignment", FormattingActions.AlignCellContentsAction));
            examples[4].Groups.Add(new SpreadsheetExample("Cell Borders", FormattingActions.AddCellBordersAction));

            // Add nodes to the "Import" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Import Arrays", ImportActions.ImportArraysAction));
            examples[5].Groups.Add(new SpreadsheetExample("Import List", ImportActions.ImportListAction));
            examples[5].Groups.Add(new SpreadsheetExample("Import Data Table", ImportActions.ImportDataTableAction));
            examples[5].Groups.Add(new SpreadsheetExample("Import Arrays with Formulas", ImportActions.ImportArrayWithFormulasAction));
            examples[5].Groups.Add(new SpreadsheetExample("Import Specified Fields from Custom Objects", ImportActions.ImportCustomObjectSpecifiedFieldsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Import Custom Objects Using Custom Converter", ImportActions.ImportCustomObjectUsingCustomConverterAction));
            
            // Add nodes to the "Export" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Export to Pdf", ExportActions.ExportToPdfAction));

            // Add nodes to the "Printing" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Print a Workbook", PrintingActions.PrintAction));

            // Add nodes to the "Document Properties" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Built-in Properties", DocumentPropertiesActions.BuiltInPropertiesAction));
            examples[8].Groups.Add(new SpreadsheetExample("Custom Properties", DocumentPropertiesActions.CustomPropertiesAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples) {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }
        

        private void button1_Click(object sender, EventArgs e) {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<Workbook> action = example.Action;
            action(workbook);
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile() {
            #region #LoadDocumentFromFile
            // Load a workbook from the file.
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }

        private void SaveDocumentToFile() {
            #region #SaveDocumentToFile
            // Save the modified document to the file.
            workbook.SaveDocument("Documents\\SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
            Process.Start("Documents\\SavedDocument.xlsx");
        }
    }
}
