using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetExamples {
    class DocumentPropertiesActions {
        #region DocumentProperties
        public static Action<IWorkbook> BuiltInPropertiesAction = BuiltInPropertiesValue;
        public static Action<IWorkbook> CustomPropertiesAction = CustomPropertiesValue;
        #endregion

        static void BuiltInPropertiesValue(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Columns[0].WidthInCharacters = 2;
                worksheet["E6"].Value = "Mike Hamilton";

                CellRange header = worksheet.Range["B2:C2"];
                header[0].Value = "Property Name";
                header[1].Value = "Value";
                header.Style = workbook.Styles[BuiltInStyleId.Accent2];

                #region #Built-inProperties
                // Set the built-in document properties.
                workbook.DocumentProperties.Title = "Spreadsheet API: document properties example";
                workbook.DocumentProperties.Description = "How to manage document properties using the Spreadsheet API";
                workbook.DocumentProperties.Keywords = "Spreadsheet, API, properties, OLEProps";
                workbook.DocumentProperties.Company = "Developer Express Inc.";

                // Display the specified built-in properties in a worksheet.
                worksheet["B3"].Value = "Title";
                worksheet["C3"].Value = workbook.DocumentProperties.Title;
                worksheet["B4"].Value = "Description";
                worksheet["C4"].Value = workbook.DocumentProperties.Description;
                worksheet["B5"].Value = "Keywords";
                worksheet["C5"].Value = workbook.DocumentProperties.Keywords;
                worksheet["B6"].Value = "Company";
                worksheet["C6"].Value = workbook.DocumentProperties.Company;
                #endregion #Built-inProperties

                worksheet.Columns.AutoFit(1, 2);
            }
            finally { workbook.EndUpdate(); }
        }

        static void CustomPropertiesValue(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Columns[0].WidthInCharacters = 2;

                CellRange header = worksheet.Range["B2:C2"];
                header[0].Value = "Property Name";
                header[1].Value = "Value";
                header.Style = workbook.Styles[BuiltInStyleId.Accent2];
                header.ColumnWidthInCharacters = 20;

                #region #CustomProperties
                // Set the custom document properties.
                workbook.DocumentProperties.Custom["Revision"] = 3;
                workbook.DocumentProperties.Custom["Completed"] = true;
                workbook.DocumentProperties.Custom["Published"] = DateTime.Now;
                #endregion #CustomProperties

                #region #LinkToContent
                //Define a name to the cell to be linked to the custom property
                workbook.DefinedNames.Add("checked_by", "E6");

                //Connect the custom property with the named cell
                workbook.DocumentProperties.Custom.LinkToContent("Checked by", "checked_by");
                #endregion #LinkToContent

                #region #DisplayCustomProperties
                // Display the specified custom properties in a worksheet.
                IEnumerable<string> customPropertiesNames = workbook.DocumentProperties.Custom.Names;
                int rowIndex = 2;
                foreach (string propertyName in customPropertiesNames)
                {
                    worksheet[rowIndex, 1].Value = propertyName;
                    worksheet[rowIndex, 2].Value = workbook.DocumentProperties.Custom[propertyName];
                    if (worksheet[rowIndex, 2].Value.IsDateTime)
                        worksheet[rowIndex, 2].NumberFormat = "[$-409]m/d/yyyy h:mm AM/PM";
                    rowIndex++;
                }
                #endregion #DisplayCustomProperties

                #region #RemoveCustomProperty
                // Remove an individual custom document property.
                workbook.DocumentProperties.Custom["Published"] = null;
                #endregion #RemoveCustomProperty

                #region #ClearCustomProperties
                // Remove all custom document properties.
                workbook.DocumentProperties.Custom.Clear();
                #endregion #ClearCustomProperties
            }
            finally { workbook.EndUpdate(); }
        }
    }
}
