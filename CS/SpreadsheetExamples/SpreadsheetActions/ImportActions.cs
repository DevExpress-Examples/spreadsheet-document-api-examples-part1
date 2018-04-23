using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Collections.Generic;
using DevExpress.Spreadsheet;

namespace SpreadsheetExamples {
    public static class ImportActions {
        #region Actions
        public static Action<Workbook> ImportArraysAction = ImportArrays;
        public static Action<Workbook> ImportListAction = ImportList;
        public static Action<Workbook> ImportDataTableAction = ImportDataTable;
        public static Action<Workbook> ImportArrayWithFormulasAction = ImportArrayWithFormulas;
        public static Action<Workbook> ImportCustomObjectSpecifiedFieldsAction = ImportCustomObjectSpecifiedFields;
        public static Action<Workbook> ImportCustomObjectUsingCustomConverterAction = ImportCustomObjectUsingCustomConverter;
        
        #endregion

        static void ImportArrays(Workbook workbook) {
            workbook.Worksheets[0].Cells["A1"].ColumnWidthInCharacters = 35;
            workbook.Worksheets[0].Cells["A1"].Value = "Import an array horizontally:";
            workbook.Worksheets[0].Cells["A3"].Value = "Import a two-dimensional array:";
            //workbook.Worksheets[0].Cells["A6"].Value = "Import data from ArrayList vertically:";
            //workbook.Worksheets[0].Cells["A11"].Value = "Import data from a DataTable:";

            #region #ImportArray
            Worksheet worksheet = workbook.Worksheets[0];
            // Create an array containing string values.
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

            // Import a two-dimensional array into the worksheet and insert it, starting with the B3 cell.
            worksheet.Import(names, 2, 1);
            #endregion #ImportTwoDimensionalArray
        }
        
        static void ImportList(Workbook workbook) {  
            #region #ImportList  
            Worksheet worksheet = workbook.Worksheets[0];
            // Create the List object containing string values.
            List<string> cities = new List<string>();
            cities.Add("New York");
            cities.Add("Rome");
            cities.Add("Beijing");
            cities.Add("Delhi");

            // Import the list into the worksheet and insert it vertically, starting with the B6 cell.
            worksheet.Import(cities, 0, 0, true);
            #endregion #ImportList
        }

        static void ImportDataTable(Workbook workbook)
        {  
            #region #ImportDataTable
            Worksheet worksheet = workbook.Worksheets[0];
            // Create the "Employees" DataTable object with four columns.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FirstN", typeof(string));
            table.Columns.Add("LastN", typeof(string));
            table.Columns.Add("JobTitle", typeof(string));
            table.Columns.Add("Age", typeof(Int32));

            table.Rows.Add("Nancy", "Davolio", "recruiter", 32);
            table.Rows.Add("Andrew", "Fuller", "engineer", 28);

            // Import data from the data table into the worksheet and insert it, starting with the B11 cell.
            worksheet.Import(table, true, 0, 0);

            // Color the table header.
            for (int i = 1; i < 5; i++) {
                worksheet.Cells[10, i].FillColor = Color.LightGray;
            }
            #endregion #ImportDataTable
        }
        
        static void ImportArrayWithFormulas(Workbook workbook)
        {
            #region  #ImportArrayWithFormulas
            Worksheet worksheet = workbook.Worksheets[0];
                
                string[] array = new string[] { "000", "=3,141", "=B1+1,01" };
                // Import data as formulas in German locale (decimal and list separators).
                worksheet.Import(array, 0, 0, false, new DataImportOptions() { ImportFormulas = true, 
                    FormulaCulture = new System.Globalization.CultureInfo("de-DE") });
                
                string[] arrayR1C1 = new string[] { "=3.141", "=R[-1]C+1.01" };
                // Import data as formulas which use R1C1 reference style.
                worksheet.Import(arrayR1C1, 1, 0, true, new DataImportOptions() { ImportFormulas = true, ReferenceStyle= ReferenceStyle.R1C1});
            #endregion  #ImportArrayWithFormulas
        }
        
        #region #ImportCustomObjectSpecifiedFields
        class MyDataObject
        {
            public MyDataObject(int intValue, string value, bool boolValue)
            {
                this.myInteger = intValue;
                this.myString = value;
                this.myBoolean = boolValue;
                
            }
                public int myInteger { get; set; }
                public string myString { get; set; }
                public bool myBoolean { get; set; }
            }

        static void ImportCustomObjectSpecifiedFields(Workbook workbook)
        {
            Worksheet worksheet = workbook.Worksheets[0];
            List<MyDataObject> list = new List<MyDataObject>();
            list.Add(new MyDataObject(1, "1", true));
            list.Add(new MyDataObject(2, "2", false));

            // Import data from the specified fields of a custom object.
            worksheet.Import(list, 0, 0, new DataSourceImportOptions() { PropertyNames = new string[] { "myBoolean", "myInteger" } });
        }
        #endregion #ImportCustomObjectSpecifiedFields
        
        #region #ImportCustomObjectUsingCustomConverter
        class MyDataValueConverter : IDataValueConverter
        {
            public bool TryConvert(object value, int columnIndex, out CellValue result)
            {
                if (columnIndex == 0)
                {
                    result = DevExpress.Docs.Text.NumberInWords.Ordinal.ConvertToText((int)value);
                    return true;
                }
                else
                    result = CellValue.FromObject(value);
                return true;
            }
        }

        static void ImportCustomObjectUsingCustomConverter(Workbook workbook)
        {
            Worksheet worksheet = workbook.Worksheets[0];
            List<MyDataObject> list = new List<MyDataObject>();
            list.Add(new MyDataObject(1, "one", true));
            list.Add(new MyDataObject(2, "two", false));
            worksheet.Import(list, 0, 0, new DataSourceImportOptions() { Converter = new MyDataValueConverter() });
        }
        #endregion #ImportCustomObjectUsingCustomConverter
    }
}
