Imports System
Imports DevExpress.Spreadsheet
Imports System.Drawing

Namespace SpreadsheetExamples

    Public Module FormulaActions

#Region "Actions"
        Public UseConstantsAndCalculationOperatorsInFormulasAction As Action(Of Workbook) = AddressOf UseConstantsAndCalculationOperatorsInFormulas

        Public R1C1ReferencesInFormulassAction As Action(Of Workbook) = AddressOf R1C1ReferencesInFormulas

        Public UseNamesInFormulasAction As Action(Of Workbook) = AddressOf UseNamesInFormulas

        Public CreateNamedFormulasAction As Action(Of Workbook) = AddressOf CreateNamedFormulas

        Public UseFunctionsInFormulasAction As Action(Of Workbook) = AddressOf UseFunctionsInFormulas

        Public CreateSharedAndArrayFormulasAction As Action(Of Workbook) = AddressOf CreateSharedAndArrayFormulas

#End Region
        Private Sub UseConstantsAndCalculationOperatorsInFormulas(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Cells("A1").Value = "Formula"
            worksheet.Cells("B1").Value = "Value"
            worksheet.Range("A1:B1").FillColor = Color.LightGray
            worksheet.Cells("A2").Value = "'= (1+5)*6"
#Region "#ConstantsAndCalculationOperators"
            ' Use constants and calculation operators in a formula.
            workbook.Worksheets(0).Cells("B2").Formula = "= (1+5)*6"
#End Region  ' #ConstantsAndCalculationOperators
        End Sub

        Private Sub R1C1ReferencesInFormulas(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Fill cells with static data.
            worksheet.Cells("A1").Value = "Data"
            worksheet.Range("A2:A11").Formula = "=ROW() - 1"
            worksheet.Cells("B1").Value = "Cell Reference Style"
            worksheet.Cells("B2").Value = "Relative R1C1 Cell Reference"
            worksheet.Cells("B3").Value = "Absolute R1C1 Cell Reference"
            worksheet.Cells("C1").Value = "Formula"
            worksheet.Cells("C2").Value = "'=SUM(RC[-3]:R[9]C[-3])"
            worksheet.Cells("C3").Value = "'=SUM(R2C1:R11C1)"
            worksheet.Cells("D1").Value = "Value"
            worksheet.Range("A1:D1").AutoFitColumns()
            worksheet.Range("A1:D1").FillColor = Color.LightGray
            worksheet.Range("A1:D11").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
#Region "#R1C1ReferencesInFormulas"
            ' Switch on the R1C1 reference style in a workbook.
            workbook.DocumentSettings.R1C1ReferenceStyle = True
            ' Specify a formula with relative R1C1 references in the "D2" cell
            ' to add values contained in cells "A2" through "A11".
            worksheet.Cells("D2").Formula = "=SUM(RC[-3]:R[9]C[-3])"
            ' Specify a formula with absolute R1C1 references 
            ' to add values contained in cells "A2" through "A11".
            worksheet.Cells("D3").Formula = "=SUM(R2C1:R11C1)"
#End Region  ' #R1C1ReferencesInFormulas
        End Sub

        Private Sub UseNamesInFormulas(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim dataRangeHeader As CellRange = worksheet.Range("A1:C1")
            worksheet.MergeCells(dataRangeHeader)
            dataRangeHeader.Value = "myRange:"
            dataRangeHeader.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            dataRangeHeader.FillColor = Color.LightGray
            worksheet.Range("A2:C5").Value = 2
            worksheet.Range("A2:C5").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            worksheet.Range("A2:C5").Borders.SetOutsideBorders(Color.LightBlue, BorderLineStyle.Medium)
            Dim sumHeader As CellRange = worksheet.Range("E1:F1")
            worksheet.MergeCells(sumHeader)
            sumHeader.Value = "Sum:"
            sumHeader.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            sumHeader.FillColor = Color.LightGray
            worksheet.Range("E2:F2").ColumnWidthInCharacters = 15
            worksheet.Cells("E2").Value = "Formula:"
            worksheet.Cells("E3").Value = "Value:"
            worksheet.Cells("F2").Value = "'= SUM(myRange)"
#Region "#NamesInFormulas"
            ' Access the "A2:C5" cell range.
            Dim range As CellRange = worksheet.Range("A2:C5")
            ' Specify the created range name.
            range.Name = "myRange"
            ' Create a formula that sums up the values of all cells included in the specified range.
            worksheet.Cells("F3").Formula = "= SUM(myRange)"
#End Region  ' #NamesInFormulas
        End Sub

        Private Sub CreateNamedFormulas(ByVal workbook As Workbook)
            workbook.Worksheets(0).Cells("A1").Value = 2
            workbook.Worksheets(0).Cells("B2").Value = 3
            workbook.Worksheets(0).Cells("C3").Value = 4
            workbook.Worksheets(1).Range("A1:C1").FillColor = Color.LightGray
            workbook.Worksheets(1).Range("A1:C1").ColumnWidthInCharacters = 25
            workbook.Worksheets(1).Cells("A1").Value = "Formula Name"
            workbook.Worksheets(1).Cells("B1").Value = "Formula"
            workbook.Worksheets(1).Cells("C1").Value = "Formula Result"
            workbook.Worksheets(1).Cells("A2").Value = "Range_Sum"
            workbook.Worksheets(1).Cells("A3").Value = "Range_DoubleSum"
            workbook.Worksheets(1).Cells("A4").Value = "-"
            workbook.Worksheets(1).Cells("B2").Value = "'=SUM(Sheet1!$A$1:$C$3)"
            workbook.Worksheets(1).Cells("B3").Value = "'=2*Sheet1!Range_Sum"
            workbook.Worksheets(1).Cells("B4").Value = "'=Range_DoubleSum + 100"
#Region "#NamedFormulas"
            Dim worksheet1 As Worksheet = workbook.Worksheets("Sheet1")
            Dim worksheet2 As Worksheet = workbook.Worksheets("Sheet2")
            ' Create a name for a formula that sums up the values of all cells
            ' included in the "A1:C3" range of the "Sheet1" worksheet. 
            ' The scope of this name is the "Sheet1" worksheet.
            worksheet1.DefinedNames.Add("Range_Sum", "=SUM(Sheet1!$A$1:$C$3)")
            ' Create a name for a formula that doubles the value resulting
            ' from the "Range_Sum" named formula.
            ' The scope of this name is the entire workbook.
            workbook.DefinedNames.Add("Range_DoubleSum", "=2*Sheet1!Range_Sum")
            ' Create formulas that use other formulas with the specified names.
            worksheet2.Cells("C2").Formula = "=Sheet1!Range_Sum"
            worksheet2.Cells("C3").Formula = "=Range_DoubleSum"
            worksheet2.Cells("C4").Formula = "=Range_DoubleSum + 100"
#End Region  ' #NamedFormulas
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets("Sheet2")
        End Sub

        Private Sub UseFunctionsInFormulas(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Fill cells with static data.
            worksheet.Cells("A1").Value = "Data"
            worksheet.Cells("A2").Value = 15
            worksheet.Range("A3:A5").Value = 3
            worksheet.Cells("A6").Value = 20
            worksheet.Cells("A7").Value = 15.12345
            worksheet.Cells("B1").ColumnWidthInCharacters = 30
            worksheet.Cells("B1").Value = "Formula String"
            worksheet.Cells("B2").Value = "'=IF(A2<10, ""Normal"", ""Excess"")"
            worksheet.Cells("B3").Value = "'=AVERAGE(A2:A7)"
            worksheet.Cells("B4").Value = "'=SUM(A3:A5,A6,100)"
            worksheet.Cells("B5").Value = "'=ROUND(SUM(A6,A7),2)"
            worksheet.Cells("B6").Value = "'=Today()"
            worksheet.Cells("B7").Value = "'=UPPER(""formula"")"
            worksheet.Cells("C1").Value = "Formula"
            worksheet.Range("A1:C1").FillColor = Color.LightGray
            worksheet.Range("A1:C7").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left
#Region "#FunctionsInFormulas"
            ' If the number in the "A2" cell is less than 10, the formula returns "Normal" 
            ' and the "C2" cell displays this text. Otherwise, the "C2" cell displays "Excess".
            worksheet.Cells("C2").Formula = "=IF(A2<10, ""Normal"", ""Excess"")"
            ' Calculate the average value for cell values within the "A2:A7" range.
            worksheet.Cells("C3").Formula = "=AVERAGE(A2:A7)"
            ' Add values contained in the "A3:A5" cell range, add the "A6" cell value, 
            ' and add 100 to that result.
            worksheet.Cells("C4").Formula = "=SUM(A3:A5,A6,100)"
            ' Use a nested function in a formula.
            ' Round the sum of the values contained in the "A6" and "A7" cells to two decimal places.
            worksheet.Cells("C5").Formula = "=ROUND(SUM(A6,A7),2)"
            ' Add the current date to the "C6" cell.
            worksheet.Cells("C6").Formula = "=Today()"
            worksheet.Cells("C6").NumberFormat = "m/d/yy"
            ' Convert the specified text to uppercase.
            worksheet.Cells("C7").Formula = "=UPPER(""formula"")"
#End Region  ' #FunctionsInFormulas
        End Sub

        Private Sub CreateSharedAndArrayFormulas(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Range("A1:D1").ColumnWidthInCharacters = 10
            worksheet.Range("A1:D1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            worksheet.Range("A1:D1").FillColor = Color.LightGray
            worksheet.MergeCells(worksheet.Range("A1:B1"))
            worksheet.Range("A1:B1").Value = "Use Shared Formulas:"
            worksheet.MergeCells(worksheet.Range("C1:D1"))
            worksheet.Range("C1:D1").Value = "Use Array Formulas:"
#Region "#SharedFormulas"
            worksheet.Cells("A2").Value = 1
            ' Use the shared formula in the "A3:A11" cell range.
            worksheet.Range("A3:A11").Formula = "=SUM(A2+1)"
            ' Use the shared formula in the "B2:B11" cell range.
            worksheet.Range("B2:B11").Formula = "=A2+2"
#End Region  ' #SharedFormulas
#Region "#ArrayFormulas"
            ' Create an array formula that multiplies values of the "A2:A11" cell range
            ' by values of the "B2:B11" cell range, 
            ' and displays the results in the "C2:C11" cell range.
            worksheet.Range("C2:C11").ArrayFormula = "=A2:A11*B2:B11"
            ' Create an array formula that multiplies values of the "C2:C11" cell range by 2
            ' and displays the results in the "D2:D11" cell range.
            worksheet.Range("D2:D11").ArrayFormula = "=C2:C11*2"
            ' Create an array formula that multiplies values of the "B2:D11" cell range, 
            ' adds the results, and displays the total sum in "D12" cell.
            worksheet.Cells("D12").ArrayFormula = "=SUM(B2:B11*C2:C11*D2:D11)"
            ' Re-dimension an array formula range:
            ' delete the array formula and create a new range with the same formula.
            If worksheet.Cells("C13").HasArrayFormula Then
                Dim af As String = worksheet.Cells("C13").ArrayFormula
                worksheet.Cells("C13").GetArrayFormulaRange().ArrayFormula = String.Empty
                worksheet.Range("C2:C11").ArrayFormula = af
            End If
#End Region  ' #ArrayFormulas
        End Sub
    End Module
End Namespace
