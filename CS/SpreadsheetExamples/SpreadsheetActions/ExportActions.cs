using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Collections.Generic;
using DevExpress.Spreadsheet;

namespace SpreadsheetExamples
{
    public static class ExportActions
    {
        #region Actions
        public static Action<Workbook> ExportToPdfAction = ExportToPdf;
        #endregion

        static void ExportToPdf(Workbook workbook)
        {
            workbook.Worksheets[0].Cells["D8"].Value = "This document is exported to the PDF format.";

            #region #ExportToPdf
            // Export the workbook to PDF.
            using (FileStream pdfFileStream = new FileStream("Documents\\Document_PDF.pdf", FileMode.Create))
            {
                workbook.ExportToPdf(pdfFileStream);
            }
            #endregion #ExportToPdf
            Process.Start("Documents\\Document_PDF.pdf");
        }
    }
}
