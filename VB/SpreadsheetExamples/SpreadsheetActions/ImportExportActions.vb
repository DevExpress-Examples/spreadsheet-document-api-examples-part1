Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet

Namespace SpreadsheetExamples
	Public NotInheritable Class ImportExportActions
		#Region "Actions"
        Public Shared ImportArraysAction As Action(Of Workbook) = AddressOf ImportArrays
        Public Shared ExportToPdfAction As Action(Of Workbook) = AddressOf ExportToPdf
		#End Region

		Private Sub New()
		End Sub
		Private Shared Sub ImportArrays(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			worksheet.Cells("A1").ColumnWidthInCharacters = 35
			worksheet.Cells("A1").Value = "Import an array horizontally:"
			worksheet.Cells("A3").Value = "Import a two-dimensional array:"
			worksheet.Cells("A6").Value = "Import data from ArrayList vertically:"
			worksheet.Cells("A11").Value = "Import data from a DataTable:"

'			#Region "#ImportArray"
			' Create the array containing string values.
			Dim array() As String = { "AAA", "BBB", "CCC", "DDD" }

			' Import the array into the worksheet and insert it horizontally, starting with the B1 cell.
			worksheet.Import(array, 0, 1, False)
'			#End Region ' #ImportArray

'			#Region "#ImportTwoDimensionalArray"
			' Create the two-dimensional array containing string values.
			Dim names(,) As String = { {"Ann", "Edward", "Angela", "Alex"}, {"Rachel", "Bruce", "Barbara", "George"} }

			' Import the two-dimensional array into the worksheet and insert it, starting with the B3 cell.
			worksheet.Import(names, 2, 1)
'			#End Region ' #ImportTwoDimensionalArray

'			#Region "#ImportList"
			' Create the List object containing string values.
			Dim cities As New List(Of String)()
			cities.Add("New York")
			cities.Add("Rome")
			cities.Add("Beijing")
			cities.Add("Delhi")

			' Import the list into the worksheet and insert it vertically, starting with the B6 cell.
			worksheet.Import(cities, 5, 1, True)
'			#End Region ' #ImportList

'			#Region "#ImportDataTable"
			' Create the "Employees" DataTable object with four columns.
			Dim table As New DataTable("Employees")
			table.Columns.Add("FirstN", GetType(String))
			table.Columns.Add("LastN", GetType(String))
			table.Columns.Add("JobTitle", GetType(String))
			table.Columns.Add("Age", GetType(Int32))

			table.Rows.Add("Nancy", "Davolio", "recruiter", 32)
			table.Rows.Add("Andrew", "Fuller", "engineer", 28)

			' Import data from the data table into the worksheet and insert it, starting with the B11 cell.
			worksheet.Import(table, True, 10, 1)

			' Color the table header.
			For i As Integer = 1 To 4
				worksheet.Cells(10, i).FillColor = Color.LightGray
			Next i
'			#End Region ' #ImportDataTable

		End Sub

		Private Shared Sub ExportToPdf(ByVal workbook As Workbook)
			workbook.Worksheets(0).Cells("D8").Value = "This document is exported to the PDF format."

'			#Region "#ExportToPdf"
			Using pdfFileStream As New FileStream("Documents\Document_PDF.pdf", FileMode.Create)
				workbook.ExportToPdf(pdfFileStream)
			End Using
'			#End Region ' #ExportToPdf
			Process.Start("Documents\Document_PDF.pdf")
		End Sub

	End Class
End Namespace
