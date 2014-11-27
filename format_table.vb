Sub FormatTable()
Dim QuellRange As Range
Set QuellRange = Selection

FirstCol = Selection.Column
FirstRow = Selection.Row

NumOfCols = Selection.Columns.Count
NumOfRows = Selection.Rows.Count

LastCol = FirstCol + NumOfCols - 1
LastRow = FirstRow + NumOfRows - 1

FormatBody Selection
FormatHeader FirstRow, FirstCol, LastCol, "Row"
FormatColumns FirstCol, LastCol, FirstRow, LastRow

QuellRange.Select

End Sub

Sub FormatTransposedTable()
Dim QuellRange As Range
Set QuellRange = Selection

FirstCol = Selection.Column
FirstRow = Selection.Row

NumOfCols = Selection.Columns.Count
NumOfRows = Selection.Rows.Count

LastCol = FirstCol + NumOfCols - 1
LastRow = FirstRow + NumOfRows - 1

FormatBody Selection
FormatHeader FirstRow, FirstCol, LastRow, "Column"
FormatColumns FirstCol, LastCol, FirstRow, LastRow

QuellRange.Select

End Sub

Sub FormatBody(Selection)

Dim xlBorders as Variant
xlBorders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
For Each xlBorder in xlBorders
	With Selection.Borders(xlBorder) ' this block draws the left border of selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With
Next xlBorder
End Sub

Sub FormatHeader(FirstRow, FirstCol, LastColOrRow, TypeOfHeader)
If TypeOfHeader = "Column" Then
	Range(ActiveSheet.Cells(FirstRow, FirstCol), ActiveSheet.Cells(LastColOrRow, FirstCol)).Select ' The header is in the row (i.e. horizontal)
Else 'Row
	Range(ActiveSheet.Cells(FirstRow, FirstCol), ActiveSheet.Cells(FirstRow, LastColOrRow)).Select ' The header is in the column (i.e. vertical)
End If

	Selection.Borders(xlDiagonalDown).LineStyle = xlNone
	Selection.Borders(xlDiagonalUp).LineStyle = xlNone

Dim xlBorders as Variant
xlBorders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

For Each xlBorder in xlBorders
	With Selection.Borders(xlBorder)
		.LineStyle = xlContinuous
		.Weight = xlMedium
		.ColorIndex = xlAutomatic
	End With
Next xlBorder

Selection.Borders(xlInsideVertical).LineStyle=xlNone

With Selection.Interior
	.ColorIndex = 15
	.Pattern = xlSolid
End With

Selection.Font.Bold = True

End Sub

Sub FormatColumns(FirstCol, LastCol, FirstRow, LastRow)

Dim xlBorders as Variant
xlBorders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

For i = FirstCol To LastCol
	Range(ActiveSheet.Cells(FirstRow, i), ActiveSheet.Cells(LastRow, i)).Select
	Selection.Borders(xlDiagonalDown).LineStyle = xlNone
	Selection.Borders(xlDiagonalUp).LineStyle = xlNone
	For Each xlBorder in xlBorders
		With Selection.Borders(xlBorder)
			.LineStyle = xlContinuous
			.ColorIndex = 0
			.Weight = xlMedium
		End With
	Next xlBorder
	Selection.Borders(xlInsideVertical).LineStyle = xlNone
	Selection.Columns.AutoFit
Next i

End Sub
