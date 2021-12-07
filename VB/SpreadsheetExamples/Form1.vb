Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetExamples

    Public Partial Class Form1
        Inherits Form

'#Region "#CreateWorkbook"
        ' Create a new Workbook object.
        Private workbook As Workbook = New Workbook()

'#End Region  ' #CreateWorkbook
        Public Sub New()
            InitializeComponent()
            InitTreeListControl()
        End Sub

        Private Sub InitTreeListControl()
            Dim examples As GroupsOfSpreadsheetExamples = New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub

        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'#Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Worksheet"))
            examples.Add(New SpreadsheetNode("Rows and Columns"))
            examples.Add(New SpreadsheetNode("Cells"))
            examples.Add(New SpreadsheetNode("Formulas"))
            examples.Add(New SpreadsheetNode("Formatting"))
            examples.Add(New SpreadsheetNode("Import/Export"))
            examples.Add(New SpreadsheetNode("Printing"))
'#End Region
'#Region "ExampleNodes"
            ' Add nodes to the "Worksheet" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Active Worksheet", AssignActiveWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("New Worksheet", AddWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Delete a Worksheet", RemoveWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Rename a Worksheet", RenameWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Copy a Worksheet within a Workbook", CopyWorksheetWithinWorkbookAction))
            examples(0).Groups.Add(New SpreadsheetExample("Copy a Worksheet between Workbooks", CopyWorksheetBetweenWorkbooksAction))
            examples(0).Groups.Add(New SpreadsheetExample("Move a Worksheet", MoveWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Show/Hide a Worksheet", ShowHideWorksheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Show/Hide Gridlines", ShowHideGridlinesAction))
            examples(0).Groups.Add(New SpreadsheetExample("Show/Hide Row and Column Headings", ShowHideHeadingsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Page Setup (View Type, Page Orientation, Page Margins, Paper Size)", PageSetupAction))
            examples(0).Groups.Add(New SpreadsheetExample("Zoom a Worksheet", ZoomWorksheetAction))
            ' Add nodes to the "Rows and Columns" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("New Row/Column", InsertRowsColumnsAction))
            examples(1).Groups.Add(New SpreadsheetExample("Delete a Row/Column", DeleteRowsColumnsAction))
            examples(1).Groups.Add(New SpreadsheetExample("Copy a Row/Column", CopyRowsColumnsAction))
            examples(1).Groups.Add(New SpreadsheetExample("Show or Hide a Row/Column", ShowHideRowsColumnsAction))
            examples(1).Groups.Add(New SpreadsheetExample("Row Height and Column Width", SpecifyRowsHeightColumnsWidthAction))
            examples(1).Groups.Add(New SpreadsheetExample("Group Rows/Columns", GroupRowsColumnsAction))
            ' Add nodes to the "Cells" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Cell Value", ChangeCellValueAction))
            examples(2).Groups.Add(New SpreadsheetExample("Named Ranges", CreateNamedRangeAction))
            examples(2).Groups.Add(New SpreadsheetExample("Add a Hyperlink to a Cell", AddHyperlinkAction))
            examples(2).Groups.Add(New SpreadsheetExample("Copy Data Only, Style Only, or Data with Style", CopyCellDataAndStyleAction))
            examples(2).Groups.Add(New SpreadsheetExample("Merge/Split Cells", MergeAndSplitCellsAction))
            examples(2).Groups.Add(New SpreadsheetExample("Clear Cells", ClearCellsAction))
            ' Add nodes to the "Formulas" group of examples. 
            examples(3).Groups.Add(New SpreadsheetExample("Constants and Calculation Operators in Formulas", UseConstantsAndCalculationOperatorsInFormulasAction))
            examples(3).Groups.Add(New SpreadsheetExample("R1C1 References in Formulas", R1C1ReferencesInFormulassAction))
            examples(3).Groups.Add(New SpreadsheetExample("Names in Formulas", UseNamesInFormulasAction))
            examples(3).Groups.Add(New SpreadsheetExample("Create Named Formulas", CreateNamedFormulasAction))
            examples(3).Groups.Add(New SpreadsheetExample("Functions in Formulas", UseFunctionsInFormulasAction))
            examples(3).Groups.Add(New SpreadsheetExample("Shared and Array Formulas", CreateSharedAndArrayFormulasAction))
            ' Add nodes to the "Formatting" group of examples.
            examples(4).Groups.Add(New SpreadsheetExample("Apply a Style", ApplayStyleAction))
            examples(4).Groups.Add(New SpreadsheetExample("Create and Modify a Style", CreateModifyStyleAction))
            examples(4).Groups.Add(New SpreadsheetExample("Individual Cell Formatting", FormatCellAction))
            examples(4).Groups.Add(New SpreadsheetExample("Date Formats", SetDateFormatsAction))
            examples(4).Groups.Add(New SpreadsheetExample("Number Formats", SetNumberFormatsAction))
            examples(4).Groups.Add(New SpreadsheetExample("Cell Colors and Background", ChangeCellColorsAction))
            examples(4).Groups.Add(New SpreadsheetExample("Font Settings", SpecifyCellFontAction))
            examples(4).Groups.Add(New SpreadsheetExample("Cell Alignment", AlignCellContentsAction))
            examples(4).Groups.Add(New SpreadsheetExample("Cell Borders", AddCellBordersAction))
            ' Add nodes to the "Import/Export" group of examples.
            examples(5).Groups.Add(New SpreadsheetExample("Import Data", ImportArraysAction))
            examples(5).Groups.Add(New SpreadsheetExample("Export to Pdf", ExportToPdfAction))
            ' Add nodes to the "Printing" group of examples.
            examples(6).Groups.Add(New SpreadsheetExample("Print a Workbook", PrintAction))
'#End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            LoadDocumentFromFile()
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then Return
            Dim action As Action(Of Workbook) = example.Action
            action(workbook)
            SaveDocumentToFile()
        End Sub

        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
'#Region "#LoadDocumentFromFile"
            ' Load a workbook from the file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
'#Region "#SaveDocumentToFile"
            ' Save the modified document to the file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #SaveDocumentToFile
            Call Process.Start("Documents\SavedDocument.xlsx")
        End Sub
    End Class
End Namespace
