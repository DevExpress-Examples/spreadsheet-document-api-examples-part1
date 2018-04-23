Imports System
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet

Namespace SpreadsheetExamples
    Public NotInheritable Class ImportActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared ImportArraysAction As Action(Of Workbook) = AddressOf ImportArrays
        Public Shared ImportListAction As Action(Of Workbook) = AddressOf ImportList
        Public Shared ImportDataTableAction As Action(Of Workbook) = AddressOf ImportDataTable
        Public Shared ImportArrayWithFormulasAction As Action(Of Workbook) = AddressOf ImportArrayWithFormulas
        Public Shared ImportCustomObjectSpecifiedFieldsAction As Action(Of Workbook) = AddressOf ImportCustomObjectSpecifiedFields
        Public Shared ImportCustomObjectUsingCustomConverterAction As Action(Of Workbook) = AddressOf ImportCustomObjectUsingCustomConverter

        #End Region

        Private Shared Sub ImportArrays(ByVal workbook As Workbook)
            workbook.Worksheets(0).Cells("A1").ColumnWidthInCharacters = 35
            workbook.Worksheets(0).Cells("A1").Value = "Import an array horizontally:"
            workbook.Worksheets(0).Cells("A3").Value = "Import a two-dimensional array:"
            'workbook.Worksheets[0].Cells["A6"].Value = "Import data from ArrayList vertically:";
            'workbook.Worksheets[0].Cells["A11"].Value = "Import data from a DataTable:";

'            #Region "#ImportArray"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create an array containing string values.
            Dim array() As String = { "AAA", "BBB", "CCC", "DDD" }

            ' Import the array into the worksheet and insert it horizontally, starting with the B1 cell.
            worksheet.Import(array, 0, 1, False)
'            #End Region ' #ImportArray

'            #Region "#ImportTwoDimensionalArray"
            ' Create the two-dimensional array containing string values.
            Dim names(,) As String = { _
                {"Ann", "Edward", "Angela", "Alex"}, _
                {"Rachel", "Bruce", "Barbara", "George"} _
            }

            ' Import a two-dimensional array into the worksheet and insert it, starting with the B3 cell.
            worksheet.Import(names, 2, 1)
'            #End Region ' #ImportTwoDimensionalArray
        End Sub

        Private Shared Sub ImportList(ByVal workbook As Workbook)
'            #Region "#ImportList  "
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create the List object containing string values.
            Dim cities As New List(Of String)()
            cities.Add("New York")
            cities.Add("Rome")
            cities.Add("Beijing")
            cities.Add("Delhi")

            ' Import the list into the worksheet and insert it vertically, starting with the B6 cell.
            worksheet.Import(cities, 0, 0, True)
'            #End Region ' #ImportList
        End Sub

        Private Shared Sub ImportDataTable(ByVal workbook As Workbook)
'            #Region "#ImportDataTable"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create the "Employees" DataTable object with four columns.
            Dim table As New DataTable("Employees")
            table.Columns.Add("FirstN", GetType(String))
            table.Columns.Add("LastN", GetType(String))
            table.Columns.Add("JobTitle", GetType(String))
            table.Columns.Add("Age", GetType(Int32))

            table.Rows.Add("Nancy", "Davolio", "recruiter", 32)
            table.Rows.Add("Andrew", "Fuller", "engineer", 28)

            ' Import data from the data table into the worksheet and insert it, starting with the B11 cell.
            worksheet.Import(table, True, 0, 0)

            ' Color the table header.
            For i As Integer = 1 To 4
                worksheet.Cells(10, i).FillColor = Color.LightGray
            Next i
'            #End Region ' #ImportDataTable
        End Sub

        Private Shared Sub ImportArrayWithFormulas(ByVal workbook As Workbook)
'            #Region " #ImportArrayWithFormulas"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

                Dim array() As String = { "000", "=3,141", "=B1+1,01" }
                ' Import data as formulas in German locale (decimal and list separators).
                worksheet.Import(array, 0, 0, False, New DataImportOptions() With {.ImportFormulas = True, .FormulaCulture = New System.Globalization.CultureInfo("de-DE")})

                Dim arrayR1C1() As String = { "=3.141", "=R[-1]C+1.01" }
                ' Import data as formulas which use R1C1 reference style.
                worksheet.Import(arrayR1C1, 1, 0, True, New DataImportOptions() With {.ImportFormulas = True, .ReferenceStyle= ReferenceStyle.R1C1})
'            #End Region '  #ImportArrayWithFormulas
        End Sub

        #Region "#ImportCustomObjectSpecifiedFields"
        Private Class MyDataObject
            Public Sub New(ByVal intValue As Integer, ByVal value As String, ByVal boolValue As Boolean)
                Me.myInteger = intValue
                Me.myString = value
                Me.myBoolean = boolValue

            End Sub
                Public Property myInteger() As Integer
                Public Property myString() As String
                Public Property myBoolean() As Boolean
        End Class

        Private Shared Sub ImportCustomObjectSpecifiedFields(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim list As New List(Of MyDataObject)()
            list.Add(New MyDataObject(1, "1", True))
            list.Add(New MyDataObject(2, "2", False))

            ' Import data from the specified fields of a custom object.
            worksheet.Import(list, 0, 0, New DataSourceImportOptions() With { _
                .PropertyNames = New String() { "myBoolean", "myInteger" } _
            })
        End Sub
        #End Region ' #ImportCustomObjectSpecifiedFields

        #Region "#ImportCustomObjectUsingCustomConverter"
        Private Class MyDataValueConverter
            Implements IDataValueConverter

            Public Function TryConvert(ByVal value As Object, ByVal columnIndex As Integer, <System.Runtime.InteropServices.Out()> ByRef result As CellValue) As Boolean Implements IDataValueConverter.TryConvert
                If columnIndex = 0 Then
                    result = DevExpress.Docs.Text.NumberInWords.Ordinal.ConvertToText(DirectCast(value, Integer))
                    Return True
                Else
                    result = CellValue.FromObject(value)
                End If
                Return True
            End Function
        End Class

        Private Shared Sub ImportCustomObjectUsingCustomConverter(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim list As New List(Of MyDataObject)()
            list.Add(New MyDataObject(1, "one", True))
            list.Add(New MyDataObject(2, "two", False))
            worksheet.Import(list, 0, 0, New DataSourceImportOptions() With {.Converter = New MyDataValueConverter()})
        End Sub
        #End Region ' #ImportCustomObjectUsingCustomConverter
    End Class
End Namespace
