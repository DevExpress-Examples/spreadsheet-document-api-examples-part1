using System;
using System.Drawing;
#region #printingUsings
using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting;
#endregion #printingUsings

namespace SpreadsheetExamples {
    public static class PrintingActions {

        public static Action<Workbook> PrintAction = Print;

        static void Print(Workbook workbook) {

            Worksheet worksheet = workbook.Worksheets[0];
 
            // Generate worksheet content - the simple multiplication table.
            Range topHeader = worksheet.Range.FromLTRB(1, 0, 20, 0);
            topHeader.Formula = "=COLUMN() - 1";
            Range leftCaption = worksheet.Range.FromLTRB(0, 1, 0, 20);
            leftCaption.Formula = "=ROW() - 1";
            Range tableRange = worksheet.Range.FromLTRB(1, 1, 20, 20);
            tableRange.Formula = "=(ROW()-1)*(COLUMN()-1)";

            

            // Format headers of the multiplication table.
            Formatting rangeFormatting = topHeader.BeginUpdateFormatting();
            rangeFormatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin;
            rangeFormatting.Borders.BottomBorder.Color = Color.Black;
            topHeader.EndUpdateFormatting(rangeFormatting);

            rangeFormatting = leftCaption.BeginUpdateFormatting();
            rangeFormatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin;
            rangeFormatting.Borders.RightBorder.Color = Color.Black;
            leftCaption.EndUpdateFormatting(rangeFormatting);

            rangeFormatting = tableRange.BeginUpdateFormatting();
            rangeFormatting.Fill.BackgroundColor = Color.LightBlue;
            tableRange.EndUpdateFormatting(rangeFormatting);

            #region #WorksheetPrintOptions
            worksheet.ActiveView.Orientation = PageOrientation.Landscape;
            //  Display row and column headings.
            worksheet.ActiveView.ShowHeadings = true;
            worksheet.ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4;
            // Access an object providing print options.
            WorksheetPrintOptions printOptions = worksheet.PrintOptions;
            //  Print in black and white.
            printOptions.BlackAndWhite = true;
            //  Do not print gridlines.
            printOptions.PrintGridlines = false;
            //  Scale the print area to fit to a  page.
            printOptions.FitToPage = true;
            //  Print a dash instead of a cell error message.
            printOptions.ErrorsPrintMode = ErrorsPrintMode.Dash;
            #endregion #WorksheetPrintOptions

            #region #PrintWorkbook
            // Invoke the Print Preview dialog for the workbook.
            using (PrintingSystem printingSystem = new PrintingSystem()) {
                using (PrintableComponentLink link = new PrintableComponentLink(printingSystem)) {
                    link.Component = workbook;
                    link.CreateDocument();
                    link.ShowPreviewDialog();
                }
            }
            #endregion #PrintWorkbook
        }
    }
}
