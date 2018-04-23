Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetExamples
	Partial Public Class Form1
		Inherits Form

		#Region "#CreateWorkbook"
		' Create a new Workbook object.
		Private workbook As New Workbook()
		#End Region ' #CreateWorkbook

		Public Sub New()
			InitializeComponent()
			InitTreeListControl()
		End Sub

		Private Sub InitTreeListControl()
			Dim examples As New GroupsOfSpreadsheetExamples()
			InitData(examples)
			DataBinding(examples)
		End Sub
		Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'			#Region "GroupNodes"
			examples.Add(New SpreadsheetNode("Worksheet"))
			examples.Add(New SpreadsheetNode("Rows and Columns"))
			examples.Add(New SpreadsheetNode("Cells"))
			examples.Add(New SpreadsheetNode("Formulas"))
			examples.Add(New SpreadsheetNode("Formatting"))
			examples.Add(New SpreadsheetNode("Import/Export"))
			examples.Add(New SpreadsheetNode("Printing"))
'			#End Region

'			#Region "ExampleNodes"
			' Add nodes to the "Worksheet" group of examples.
			examples(0).Groups.Add(New SpreadsheetExample("Active Worksheet", WorksheetActions.AssignActiveWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("New Worksheet", WorksheetActions.AddWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("Delete a Worksheet", WorksheetActions.RemoveWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("Rename a Worksheet", WorksheetActions.RenameWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("Copy a Worksheet within a Workbook", WorksheetActions.CopyWorksheetWithinWorkbookAction))
			examples(0).Groups.Add(New SpreadsheetExample("Copy a Worksheet between Workbooks", WorksheetActions.CopyWorksheetBetweenWorkbooksAction))
			examples(0).Groups.Add(New SpreadsheetExample("Move a Worksheet", WorksheetActions.MoveWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("Show/Hide a Worksheet", WorksheetActions.ShowHideWorksheetAction))
			examples(0).Groups.Add(New SpreadsheetExample("Show/Hide Gridlines", WorksheetActions.ShowHideGridlinesAction))
			examples(0).Groups.Add(New SpreadsheetExample("Show/Hide Row and Column Headings", WorksheetActions.ShowHideHeadingsAction))
			examples(0).Groups.Add(New SpreadsheetExample("Page Setup (View Type, Page Orientation, Page Margins, Paper Size)", WorksheetActions.PageSetupAction))
			examples(0).Groups.Add(New SpreadsheetExample("Zoom a Worksheet", WorksheetActions.ZoomWorksheetAction))

			' Add nodes to the "Rows and Columns" group of examples.
			examples(1).Groups.Add(New SpreadsheetExample("New Row/Column", RowAndColumnActions.InsertRowsColumnsAction))
			examples(1).Groups.Add(New SpreadsheetExample("Delete a Row/Column", RowAndColumnActions.DeleteRowsColumnsAction))
			examples(1).Groups.Add(New SpreadsheetExample("Copy a Row/Column", RowAndColumnActions.CopyRowsColumnsAction))
			examples(1).Groups.Add(New SpreadsheetExample("Show or Hide a Row/Column", RowAndColumnActions.ShowHideRowsColumnsAction))
			examples(1).Groups.Add(New SpreadsheetExample("Row Height and Column Width", RowAndColumnActions.SpecifyRowsHeightColumnsWidthAction))
			examples(1).Groups.Add(New SpreadsheetExample("Group Rows/Columns", RowAndColumnActions.GroupRowsColumnsAction))

			' Add nodes to the "Cells" group of examples.
			examples(2).Groups.Add(New SpreadsheetExample("Cell Value", CellActions.ChangeCellValueAction))
			examples(2).Groups.Add(New SpreadsheetExample("Named Ranges", CellActions.CreateNamedRangeAction))
			examples(2).Groups.Add(New SpreadsheetExample("Add a Hyperlink to a Cell", CellActions.AddHyperlinkAction))
			examples(2).Groups.Add(New SpreadsheetExample("Copy Data Only, Style Only, or Data with Style", CellActions.CopyCellDataAndStyleAction))
			examples(2).Groups.Add(New SpreadsheetExample("Merge/Split Cells", CellActions.MergeAndSplitCellsAction))
			examples(2).Groups.Add(New SpreadsheetExample("Clear Cells", CellActions.ClearCellsAction))

			' Add nodes to the "Formulas" group of examples. 
			examples(3).Groups.Add(New SpreadsheetExample("Constants and Calculation Operators in Formulas", FormulaActions.UseConstantsAndCalculationOperatorsInFormulasAction))
			examples(3).Groups.Add(New SpreadsheetExample("R1C1 References in Formulas", FormulaActions.R1C1ReferencesInFormulassAction))
			examples(3).Groups.Add(New SpreadsheetExample("Names in Formulas", FormulaActions.UseNamesInFormulasAction))
			examples(3).Groups.Add(New SpreadsheetExample("Create Named Formulas", FormulaActions.CreateNamedFormulasAction))
			examples(3).Groups.Add(New SpreadsheetExample("Functions in Formulas", FormulaActions.UseFunctionsInFormulasAction))
			examples(3).Groups.Add(New SpreadsheetExample("Shared and Array Formulas", FormulaActions.CreateSharedAndArrayFormulasAction))

			' Add nodes to the "Formatting" group of examples.
			examples(4).Groups.Add(New SpreadsheetExample("Apply a Style", FormattingActions.ApplayStyleAction))
			examples(4).Groups.Add(New SpreadsheetExample("Create and Modify a Style", FormattingActions.CreateModifyStyleAction))
			examples(4).Groups.Add(New SpreadsheetExample("Individual Cell Formatting", FormattingActions.FormatCellAction))
			examples(4).Groups.Add(New SpreadsheetExample("Date Formats", FormattingActions.SetDateFormatsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Number Formats", FormattingActions.SetNumberFormatsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Cell Colors and Background", FormattingActions.ChangeCellColorsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Font Settings", FormattingActions.SpecifyCellFontAction))
			examples(4).Groups.Add(New SpreadsheetExample("Cell Alignment", FormattingActions.AlignCellContentsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Cell Borders", FormattingActions.AddCellBordersAction))

			' Add nodes to the "Import/Export" group of examples.
			examples(5).Groups.Add(New SpreadsheetExample("Import Data", ImportExportActions.ImportArraysAction))
			examples(5).Groups.Add(New SpreadsheetExample("Export to Pdf", ImportExportActions.ExportToPdfAction))

			' Add nodes to the "Printing" group of examples.
			examples(6).Groups.Add(New SpreadsheetExample("Print a Workbook", PrintingActions.PrintAction))
'			#End Region
		End Sub

		Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
			treeList1.DataSource = examples
			treeList1.ExpandAll()
			treeList1.BestFitColumns()
		End Sub


		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			LoadDocumentFromFile()
			Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
			If example Is Nothing Then
				Return
			End If
			Dim action As Action(Of Workbook) = example.Action
			action(workbook)
			SaveDocumentToFile()
		End Sub

		' ------------------- Load and Save a Document -------------------
		Private Sub LoadDocumentFromFile()
'			#Region "#LoadDocumentFromFile"
			' Load a workbook from the file.
			workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #LoadDocumentFromFile
		End Sub

		Private Sub SaveDocumentToFile()
'			#Region "#SaveDocumentToFile"
			' Save the modified document to the file.
			workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #SaveDocumentToFile
			Process.Start("Documents\SavedDocument.xlsx")
		End Sub
	End Class
End Namespace
