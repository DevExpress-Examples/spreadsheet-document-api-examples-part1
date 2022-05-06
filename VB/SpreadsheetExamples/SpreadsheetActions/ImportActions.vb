Imports System
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet
Imports System.Runtime.InteropServices

Namespace SpreadsheetExamples

    Public Module ImportActions

#Region "Actions"
        Public ImportArraysAction As Action(Of Workbook) = AddressOf ImportArrays

        Public ImportListAction As Action(Of Workbook) = AddressOf ImportList

        Public ImportDataTableAction As Action(Of Workbook) = AddressOf ImportDataTable

        Public ImportArrayWithFormulasAction As Action(Of Workbook) = AddressOf ImportArrayWithFormulas

        Public ImportCustomObjectSpecifiedFieldsAction As Action(Of Workbook) = AddressOf ImportCustomObjectSpecifiedFields

        Public ImportCustomObjectUsingCustomConverterAction As Action(Of Workbook) = AddressOf ImportCustomObjectUsingCustomConverter

#End Region
        Private Sub ImportArrays(ByVal workbook As Workbook)
            workbook.Worksheets(0).Cells("A1").ColumnWidthInCharacters = 35
            workbook.Worksheets(0).Cells("A1").Value = "Import an array horizontally:"
            workbook.Worksheets(0).Cells("A3").Value = "Import a two-dimensional array:"
#Region "#ImportArray"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create an array of strings.
            Dim array As String() = New String() {"AAA", "BBB", "CCC", "DDD"}
            ' Insert array values into the worksheet horizontally.
            ' Data import starts with the "B1" cell.
            worksheet.Import(array, 0, 1, False)
#End Region  ' #ImportArray
#Region "#ImportTwoDimensionalArray"
            ' Create a two-dimensional array of strings.
            Dim names As String(,) = New String(1, 3) {{"Ann", "Edward", "Angela", "Alex"}, {"Rachel", "Bruce", "Barbara", "George"}}
            ' Insert array values into the worksheet.
            ' Data import starts with the "B3" cell.
            worksheet.Import(names, 2, 1)
#End Region  ' #ImportTwoDimensionalArray
        End Sub

        Private Sub ImportList(ByVal workbook As Workbook)
#Region "#ImportList  "
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create a list that contains string values.
            Dim cities As List(Of String) = New List(Of String)()
            cities.Add("New York")
            cities.Add("Rome")
            cities.Add("Beijing")
            cities.Add("Delhi")
            ' Insert list values into the worksheet vertically.
            ' Data import starts with the "A1" cell.
            worksheet.Import(cities, 0, 0, True)
#End Region  ' #ImportList
        End Sub

        Private Sub ImportDataTable(ByVal workbook As Workbook)
#Region "#ImportDataTable"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create an "Employees" DataTable object with four columns.
            Dim table As DataTable = New DataTable("Employees")
            table.Columns.Add("FirstN", GetType(String))
            table.Columns.Add("LastN", GetType(String))
            table.Columns.Add("JobTitle", GetType(String))
            table.Columns.Add("Age", GetType(Integer))
            table.Rows.Add("Nancy", "Davolio", "recruiter", 32)
            table.Rows.Add("Andrew", "Fuller", "engineer", 28)
            ' Insert data table values into the worksheet.
            ' Data import starts with the "A1" cell.
            worksheet.Import(table, True, 0, 0)
            ' Color the table header.
            For i As Integer = 1 To 5 - 1
                worksheet.Cells(10, i).FillColor = Color.LightGray
            Next
#End Region  ' #ImportDataTable
        End Sub

        Private Sub ImportArrayWithFormulas(ByVal workbook As Workbook)
#Region "#ImportArrayWithFormulas"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim array As String() = New String() {"000", "=3,141", "=B1+1,01"}
            ' Import data as formulas in German locale (decimal and list separators).
            worksheet.Import(array, 0, 0, False, New DataImportOptions() With {.ImportFormulas = True, .FormulaCulture = New Globalization.CultureInfo("de-DE")})
            Dim arrayR1C1 As String() = New String() {"=3.141", "=R[-1]C+1.01"}
            ' Import data as formulas that use the R1C1 reference style.
            worksheet.Import(arrayR1C1, 1, 0, True, New DataImportOptions() With {.ImportFormulas = True, .ReferenceStyle = ReferenceStyle.R1C1})
#End Region  ' #ImportArrayWithFormulas
        End Sub

#Region "#ImportCustomObjectSpecifiedFields"
        Private Class MyDataObject

            Public Sub New(ByVal intValue As Integer, ByVal value As String, ByVal boolValue As Boolean)
                myInteger = intValue
                myString = value
                myBoolean = boolValue
            End Sub

            Public Property myInteger As Integer

            Public Property myString As String

            Public Property myBoolean As Boolean
        End Class

        Private Sub ImportCustomObjectSpecifiedFields(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim list As List(Of MyDataObject) = New List(Of MyDataObject)()
            list.Add(New MyDataObject(1, "1", True))
            list.Add(New MyDataObject(2, "2", False))
            ' Import values from specific data source fields.
            worksheet.Import(list, 0, 0, New DataSourceImportOptions() With {.PropertyNames = New String() {"myBoolean", "myInteger"}})
        End Sub

#End Region  ' #ImportCustomObjectSpecifiedFields
#Region "#ImportCustomObjectUsingCustomConverter"
        ' A custom converter that converts the first column's integer values to text.
        Private Class MyDataValueConverter
            Implements IDataValueConverter

            Public Function TryConvert(ByVal value As Object, ByVal columnIndex As Integer, <Out> ByRef result As CellValue) As Boolean Implements IDataValueConverter.TryConvert
                If columnIndex = 0 Then
                    result = DevExpress.Docs.Text.NumberInWords.Ordinal.ConvertToText(CInt(value))
                    Return True
                Else
                    result = CellValue.FromObject(value)
                End If

                Return True
            End Function
        End Class

        Private Sub ImportCustomObjectUsingCustomConverter(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim list As List(Of MyDataObject) = New List(Of MyDataObject)()
            list.Add(New MyDataObject(1, "one", True))
            list.Add(New MyDataObject(2, "two", False))
            ' Import values from specific data source fields and converts integer values to text.
            worksheet.Import(list, 0, 0, New DataSourceImportOptions() With {.Converter = New MyDataValueConverter()})
        End Sub
#End Region  ' #ImportCustomObjectUsingCustomConverter
    End Module
End Namespace
