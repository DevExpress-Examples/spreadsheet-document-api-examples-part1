Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetExamples
	Public NotInheritable Class CellActions

		#Region "Actions"
        Public Shared ChangeCellValueAction As Action(Of Workbook) = AddressOf ChangeCellValue
        Public Shared CreateNamedRangeAction As Action(Of Workbook) = AddressOf CreateNamedRange
        Public Shared AddHyperlinkAction As Action(Of Workbook) = AddressOf AddHyperlink
        Public Shared CopyCellDataAndStyleAction As Action(Of Workbook) = AddressOf CopyCellDataAndStyle
        Public Shared MergeAndSplitCellsAction As Action(Of Workbook) = AddressOf MergeAndSplitCells
        Public Shared ClearCellsAction As Action(Of Workbook) = AddressOf ClearCells
		#End Region

		Private Sub New()
		End Sub
		Private Shared Sub ChangeCellValue(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			worksheet.Cells("A1").Value = "dateTime:"
			worksheet.Cells("A2").Value = "double:"
			worksheet.Cells("A3").Value = "string:"
			worksheet.Cells("A4").Value = "error constant:"
			worksheet.Cells("A5").Value = "boolean:"
			worksheet.Cells("A6").Value = "float:"
			worksheet.Cells("A7").Value = "char:"
			worksheet.Cells("A8").Value = "int32:"
			worksheet.Cells("A10").Value = "Fill a range of cells:"

			worksheet.Columns("A").WidthInCharacters = 20
			worksheet.Columns("B").WidthInCharacters = 20
			worksheet.Range("A1:B8").Alignment.Horizontal = HorizontalAlignment.Left

'			#Region "#CellValue"
			' Add data of different types to cells.
			worksheet.Cells("B1").Value = DateTime.Now
			worksheet.Cells("B2").Value = Math.PI
			worksheet.Cells("B3").Value = "Have a nice day!"
			worksheet.Cells("B4").Value = CellValue.ErrorReference
			worksheet.Cells("B5").Value = True
			worksheet.Cells("B6").Value = Single.MaxValue
			worksheet.Cells("B7").Value = "a"c
			worksheet.Cells("B8").Value = Int32.MaxValue

			' Fill all cells in the range with 10.
			worksheet.Range("B10:E10").Value = 10
'			#End Region ' #CellValue
		End Sub

		Private Shared Sub CreateNamedRange(ByVal workbook As Workbook)

'			#Region "#NamedRange"
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Create a range.
			Dim rangeB3D6 As Range = worksheet.Range("B3:D6")
			' Specify the name for the created range.
			rangeB3D6.Name = "rangeB3D6"

			' Create a new defined name with the specifed range name and absolute reference.
			Dim definedName As DefinedName = worksheet.DefinedNames.Add("rangeB17D20", "Sheet1!$B$17:$D$20")
			' Create a range using the specified defined name.
			Dim B17D20 As Range = worksheet.Range(definedName.Name)
'			#End Region ' #NamedRange
		End Sub

		Private Shared Sub AddHyperlink(ByVal workbook As Workbook)

			Dim worksheet As Worksheet = workbook.Worksheets(0)
			worksheet.Range("A:C").ColumnWidthInCharacters = 12

'			#Region "#AddHyperlink"
			' Create a hyperlink to a web page.
			Dim cell As Cell = worksheet.Cells("A1")
			worksheet.Hyperlinks.Add(cell, "http://www.devexpress.com/", True, "DevExpress")

			' Create a hyperlink to a cell range in a workbook.
			Dim range As Range = worksheet.Range("C3:D4")
			Dim cellHyperlink As Hyperlink = worksheet.Hyperlinks.Add(range, "Sheet2!B2:E7", False, "Select Range")
			cellHyperlink.TooltipText = "Click Me"
'			#End Region ' #AddHyperlink
		End Sub

		Private Shared Sub CopyCellDataAndStyle(ByVal workbook As Workbook)
'			#Region "#CopyCell"
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			worksheet.Columns("A").WidthInCharacters = 32
			worksheet.Columns("B").WidthInCharacters = 20
			Dim style As Style = workbook.Styles(BuiltInStyleId.Input)

			' Specify the content and formatting for a source cell.
			worksheet.Cells("A1").Value = "Source Cell"

			Dim sourceCell As Cell = worksheet.Cells("B1")
			sourceCell.Formula = "= PI()"
			sourceCell.NumberFormat = "0.0000"
			sourceCell.Style = style
			sourceCell.Font.Color = Color.Blue
			sourceCell.Font.Bold = True
			sourceCell.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin)

			' Copy all information from the source cell to the "B3" cell. 
			worksheet.Cells("A3").Value = "Copy All"
			worksheet.Cells("B3").CopyFrom(sourceCell)

			' Copy only the source cell content (e.g., text, numbers, formula calculated values) to the "B4" cell.
			worksheet.Cells("A4").Value = "Copy Values"
			worksheet.Cells("B4").CopyFrom(sourceCell, PasteSpecial.Values)

			' Copy the source cell content (e.g., text, numbers, formula calculated values) 
			' and number formats to the "B5" cell.
			worksheet.Cells("A5").Value = "Copy Values and Number Formats"
			worksheet.Cells("B5").CopyFrom(sourceCell, PasteSpecial.Values Or PasteSpecial.NumberFormats)

			' Copy only the formatting information from the source cell to the "B6" cell.
			worksheet.Cells("A6").Value = "Copy Formats"
			worksheet.Cells("B6").CopyFrom(sourceCell, PasteSpecial.Formats)

			' Copy all information from the source cell to the "B7" cell except for border settings.
			worksheet.Cells("A7").Value = "Copy All Except Borders"
			worksheet.Cells("B7").CopyFrom(sourceCell, PasteSpecial.All And (Not PasteSpecial.Borders))

			' Copy information only about borders from the source cell to the "B8" cell.
			worksheet.Cells("A8").Value = "Copy Borders"
			worksheet.Cells("B8").CopyFrom(sourceCell, PasteSpecial.Borders)
'			#End Region ' #CopyCell
		End Sub

		Private Shared Sub MergeAndSplitCells(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			worksheet.Cells("A1").FillColor = Color.LightGray

			worksheet.Cells("B2").Value = "B2"
			worksheet.Cells("B2").FillColor = Color.LightGreen

			worksheet.Cells("C3").Value = "C3"
			worksheet.Cells("C3").FillColor = Color.LightSalmon

'			#Region "#MergeCells"
			' Merge cells contained in the range.
			worksheet.MergeCells(worksheet.Range("A1:C5"))
'			#End Region ' #MergeCells
		End Sub

		Private Shared Sub ClearCells(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			worksheet.Range("A:D").ColumnWidthInCharacters = 30
			worksheet.Range("B1:D6").Alignment.Horizontal = DevExpress.Spreadsheet.HorizontalAlignment.Center

			worksheet("B1").Value = "Initial Cell Content and Formatting:"			
			worksheet.MergeCells(worksheet("C1:D1"))			
			worksheet("C1:D1").Value = "Cleared Cells:"

			worksheet("A2").Value = "Clear All:"
			worksheet("A3").Value = "Clear Cell Content Only:"
			worksheet("A4").Value = "Clear Cell Formatting Only:"
			worksheet("A5").Value = "Clear Cell Hyperlinks Only:"
			worksheet("A6").Value = "Clear Cell Comments Only:"

			' Specify initial content and formatting for cells.
			Dim sourceCells As Range = worksheet("B2:D6")
			sourceCells.Value = DateTime.Now
			sourceCells.Style = workbook.Styles(BuiltInStyleId.Accent3_40percent)
			sourceCells.Font.Color = Color.LightSeaGreen
			sourceCells.Font.Bold = True
			sourceCells.Borders.SetAllBorders(Color.Blue, BorderLineStyle.Dashed)
			worksheet.Hyperlinks.Add(worksheet("B5"), "http://www.devexpress.com/", True, "DevExpress")
			worksheet.Hyperlinks.Add(worksheet("C5"), "http://www.devexpress.com/", True, "DevExpress")
			worksheet.Hyperlinks.Add(worksheet("D5"), "http://www.devexpress.com/", True, "DevExpress")
			worksheet.Comments.Add(worksheet("B6"), "Me", "Cell Comment")
			worksheet.Comments.Add(worksheet("C6"), "Me", "Cell Comment")
			worksheet.Comments.Add(worksheet("D6"), "Me", "Cell Comment")

'			#Region "#ClearCell"
			 ' Remove all cell information (content, formatting, hyperlinks and comments).
			worksheet.Clear(worksheet("C2:D2"))

			' Remove cell content.
			worksheet.ClearContents(worksheet("C3"))
			worksheet("D3").Value = Nothing

			' Remove cell formatting.
			worksheet.ClearFormats(worksheet("C4"))
			worksheet("D4").Style = workbook.Styles.DefaultStyle

			' Remove hyperlinks from cells.
			worksheet.ClearHyperlinks(worksheet("C5"))

			Dim hyperlinkD5 As Hyperlink = worksheet.Hyperlinks.GetHyperlinks(worksheet("D5"))(0)
			worksheet.Hyperlinks.Remove(hyperlinkD5)

			' Remove comments from cells.
			worksheet.ClearComments(worksheet("C6"))

			Dim commentD6 As Comment = worksheet.Comments.GetComments(worksheet("D6"))(0)
			worksheet.Comments.Remove(commentD6)
'			#End Region ' #ClearCell
		End Sub
	End Class
End Namespace
