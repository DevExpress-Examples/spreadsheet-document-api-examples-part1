using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Collections.Generic;
using DevExpress.Spreadsheet;

namespace SpreadsheetExamples {
    public static class ImportExportActions {
        #region Actions
        public static Action<Workbook> ImportArraysAction = ImportArrays;
        public static Action<Workbook> ExportToPdfAction = ExportToPdf;
        #endregion

        static void ImportArrays(Workbook workbook) {   
            Worksheet worksheet = workbook.Worksheets[0];
            
            worksheet.Cells["A1"].ColumnWidthInCharacters = 35;
            worksheet.Cells["A1"].Value = "Import an array horizontally:";
            worksheet.Cells["A3"].Value = "Import a two-dimensional array:";
            worksheet.Cells["A6"].Value = "Import data from ArrayList vertically:";
            worksheet.Cells["A11"].Value = "Import data from a DataTable:";

            #region #ImportArray
            // Create the array containing string values.
            string[] array = new string[] { "AAA", "BBB", "CCC", "DDD" };

            // Import the array into the worksheet and insert it horizontally, starting with the B1 cell.
            worksheet.Import(array, 0, 1, false);
            #endregion #ImportArray

            #region #ImportTwoDimensionalArray
            // Create the two-dimensional array containing string values.
            String[,] names = new String[2, 4]{
            {"Ann", "Edward", "Angela", "Alex"},
            {"Rachel", "Bruce", "Barbara", "George"}
                 };

            // Import the two-dimensional array into the worksheet and insert it, starting with the B3 cell.
            worksheet.Import(names, 2, 1);
            #endregion #ImportTwoDimensionalArray

            #region #ImportList
            // Create the List object containing string values.
            List<string> cities = new List<string>();
            cities.Add("New York");
            cities.Add("Rome");
            cities.Add("Beijing");
            cities.Add("Delhi");

            // Import the list into the worksheet and insert it vertically, starting with the B6 cell.
            worksheet.Import(cities, 5, 1, true);
            #endregion #ImportList

            #region #ImportDataTable
            // Create the "Employees" DataTable object with four columns.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FirstN", typeof(string));
            table.Columns.Add("LastN", typeof(string));
            table.Columns.Add("JobTitle", typeof(string));
            table.Columns.Add("Age", typeof(Int32));

            table.Rows.Add("Nancy", "Davolio", "recruiter", 32);
            table.Rows.Add("Andrew", "Fuller", "engineer", 28);

            // Import data from the data table into the worksheet and insert it, starting with the B11 cell.
            worksheet.Import(table, true, 10, 1);

            // Color the table header.
            for (int i = 1; i < 5; i++) {
                worksheet.Cells[10, i].FillColor = Color.LightGray;
            }
            #endregion #ImportDataTable

        }

        static void ExportToPdf(Workbook workbook) {
            workbook.Worksheets[0].Cells["D8"].Value = "This document is exported to the PDF format.";

            #region #ExportToPdf
            using (FileStream pdfFileStream = new FileStream("Documents\\Document_PDF.pdf", FileMode.Create)) {
                workbook.ExportToPdf(pdfFileStream);
            }
            #endregion #ExportToPdf
            Process.Start("Documents\\Document_PDF.pdf");
        }

    }
}
