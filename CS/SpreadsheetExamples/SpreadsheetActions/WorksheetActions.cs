using System;
using System.Drawing;
using DevExpress.Spreadsheet;

namespace SpreadsheetExamples {

    public static class WorksheetActions {
        #region Actions
        public static Action<Workbook> AssignActiveWorksheetAction = AssignActiveWorksheet;
        public static Action<Workbook> AddWorksheetAction = AddWorksheet;
        public static Action<Workbook> RemoveWorksheetAction = RemoveWorksheet;
        public static Action<Workbook> RenameWorksheetAction = RenameWorksheet;
        public static Action<Workbook> CopyWorksheetWithinWorkbookAction = CopyWorksheetWithinWorkbook;
        public static Action<Workbook> CopyWorksheetBetweenWorkbooksAction = CopyWorksheetBetweenWorkbooks;
        public static Action<Workbook> MoveWorksheetAction = MoveWorksheet;
        public static Action<Workbook> ShowHideWorksheetAction = ShowHideWorksheet;
        public static Action<Workbook> ShowHideGridlinesAction = ShowHideGridlines;
        public static Action<Workbook> ShowHideHeadingsAction = ShowHideHeadings;
        public static Action<Workbook> PageSetupAction = PageSetup;
        public static Action<Workbook> ZoomWorksheetAction = ZoomWorksheet;
        
        #endregion

        static void AssignActiveWorksheet(Workbook workbook) {
            #region #ActiveWorksheet
            // Set the second worksheet under the "Sheet2" name as active.
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["Sheet2"];
            #endregion #ActiveWorksheet
        }

        static void AddWorksheet(Workbook workbook) {
            #region #AddWorksheet
            // Add a new worksheet to the workbook. The worksheet will be inserted into the end of the existing worksheet collection
            // under the name "SheetN", where N is a number following the largest number used in worksheet names in the previously existing collection.
            workbook.Worksheets.Add();

            // Add a new worksheet under the specified name.
            workbook.Worksheets.Add().Name = "TestSheet1";

            workbook.Worksheets.Add("TestSheet2");

            // Add a new worksheet to the specified position in the collection of worksheets.
            workbook.Worksheets.Insert(1, "TestSheet3");

            workbook.Worksheets.Insert(3);

            #endregion #AddWorksheet
        }

        static void RemoveWorksheet(Workbook workbook) {
            #region #DeleteWorksheet
            // Delete the "Sheet2" worksheet from the workbook.
            workbook.Worksheets.Remove(workbook.Worksheets["Sheet2"]);

            // Delete the first worksheet from the workbook.
            workbook.Worksheets.RemoveAt(0);
            #endregion #DeleteWorksheet
        }

        static void RenameWorksheet(Workbook workbook) {
            #region #RenameWorksheet
            // Change the name of the second worksheet.
            workbook.Worksheets[1].Name = "Renamed Sheet";
            #endregion #RenameWorksheet
        }

        static void CopyWorksheetWithinWorkbook(Workbook workbook) {

            workbook.Worksheets["Sheet1"].Cells.FillColor = Color.LightSteelBlue;
            workbook.Worksheets["Sheet1"].Cells["A1"].ColumnWidthInCharacters = 50;
            workbook.Worksheets["Sheet1"].Cells["A1"].Value = "Sheet1's Content";

            #region #CopyWorksheet
            // Add a new worksheet to a workbook.
            workbook.Worksheets.Add("Sheet1_Copy");

            // Copy all information (content and formatting) to the newly created worksheet 
            // from the "Sheet1" worksheet.
            workbook.Worksheets["Sheet1_Copy"].CopyFrom(workbook.Worksheets["Sheet1"]);
            #endregion #CopyWorksheet
        }

        static void CopyWorksheetBetweenWorkbooks(Workbook workbook)
        {
            #region #CopyWorksheetsBetweenWorkbooks
            // Create a source workbook.
            Workbook sourceWorkbook = new Workbook();

            // Add a new worksheet.
            sourceWorkbook.Worksheets.Add();

            // Modify the second worksheet of the source workbook to be copied.
            sourceWorkbook.Worksheets[1].Cells["A1"].Value = "A worksheet to be copied";
            sourceWorkbook.Worksheets[1].Cells["A1"].Font.Color = Color.ForestGreen;

            // Copy the second worksheet of the source workbook into the first worksheet of another workbook.
            workbook.Worksheets[0].CopyFrom(sourceWorkbook.Worksheets[1]);
            #endregion #CopyWorksheetsBetweenWorkbooks
        }

        static void MoveWorksheet(Workbook workbook) {
            #region #MoveWorksheet
            // Move the first worksheet to the position of the last worksheet within the workbook.
            workbook.Worksheets[0].Move(workbook.Worksheets.Count - 1);
            #endregion #MoveWorksheet
        }

        static void ShowHideWorksheet(Workbook workbook) {
            #region #ShowHideWorksheet
            // Hide the worksheet under the "Sheet2" name and prevent end-users from unhiding it via user interface.
            // To make this worksheet visible again, use the Worksheet.Visible property.
            workbook.Worksheets["Sheet2"].VisibilityType = WorksheetVisibilityType.VeryHidden;

            // Hide the "Sheet3" worksheet. 
            // In this state a worksheet can be unhidden via user interface.
            workbook.Worksheets["Sheet3"].Visible = false;
            #endregion #ShowHideWorksheet
        }

        static void ShowHideGridlines(Workbook workbook) {
            #region #ShowHideGridlines
            // Hide gridlines on the first worksheet.
            workbook.Worksheets[0].ActiveView.ShowGridlines = false;
            #endregion #ShowHideGridlines
        }

        static void ShowHideHeadings(Workbook workbook) {
            #region #ShowHideHeadings
            // Hide row and column headings in the first worksheet.
            workbook.Worksheets[0].ActiveView.ShowHeadings = false;
            #endregion #ShowHideHeadings
        }

        static void PageSetup(Workbook workbook)
        {
            #region #ViewType
            // Select the worksheet view type.
            workbook.Worksheets[0].ActiveView.ViewType = WorksheetViewType.PageLayout;
            #endregion #ViewType

            #region #PageOrientation
            // Set the page orientation to Landscape.
            workbook.Worksheets[0].ActiveView.Orientation = PageOrientation.Landscape;
            #endregion #PageOrientation

            #region #PageMargins
            // Select a unit of measure used within the workbook.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Access page margins.
            Margins pageMargins = workbook.Worksheets[0].ActiveView.Margins;

            // Specify page margins.
            pageMargins.Left = 1;
            pageMargins.Top = 1.5F;
            pageMargins.Right = 1;
            pageMargins.Bottom = 0.8F;

            // Specify header and footer margins.
            pageMargins.Header = 1;
            pageMargins.Footer = 0.4F;
            #endregion #PageMargins

            #region #PaperSize
            // Select the page's paper size.
            workbook.Worksheets[0].ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4;
            #endregion #PaperSize
        }

        static void ZoomWorksheet(Workbook workbook) {
            #region #WorksheetZoom
            // Zoom in the worksheet view. 
            workbook.Worksheets[0].ActiveView.Zoom = 150;            
            #endregion #WorksheetZoom
        }

    }
}
