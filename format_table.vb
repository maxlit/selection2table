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
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone

	With Selection.Borders(xlEdgeLeft) ' this block draws the left border of selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeTop) ' this block draws the top border of selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeBottom) ' this block draws the bottom border of selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeRight) ' this block draws the right border of selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlInsideVertical) ' this block puts the vertical lines inside selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlInsideHorizontal) ' this block puts the horizontal lines inside selection
		.LineStyle = xlContinuous
		.Weight = xlThin
		.ColorIndex = xlAutomatic
	End With

End Sub

Sub FormatHeader(FirstRow, FirstCol, LastColOrRow, TypeOfHeader)
If TypeOfHeader = "Column" Then
	Range(ActiveSheet.Cells(FirstRow, FirstCol), ActiveSheet.Cells(LastRowOrCol, FirstCol)).Select ' The header is in the row (i.e. horizontal)
Else 'Row
	Range(ActiveSheet.Cells(FirstRow, FirstCol), ActiveSheet.Cells(FirstRow, LastRowOrCol)).Select ' The header is in the column (i.e. vertical)
End If

	Selection.Borders(xlDiagonalDown).LineStyle = xlNone
	Selection.Borders(xlDiagonalUp).LineStyle = xlNone

	With Selection.Borders(xlEdgeLeft)
		.LineStyle = xlContinuous
		.Weight = xlMedium
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeTop)
		.LineStyle = xlContinuous
		.Weight = xlMedium
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeBottom)
		.LineStyle = xlContinuous
		.Weight = xlMedium
		.ColorIndex = xlAutomatic
	End With

	With Selection.Borders(xlEdgeRight)
		.LineStyle = xlContinuous
		.Weight = xlMedium
		.ColorIndex = xlAutomatic
	End With

	Selection.Borders(xlInsideVertical).LineStyle=xlNone

	With Selection.Interior
		.ColorIndex = 15
		.Pattern = xlSolid
	End With
	Selection.Font.Bold = True

End Sub

Sub FormatColumns(FirstCol, LastCol, FirstRow, LastRow)
For i = FirstCol To LastCol
	Range(ActiveSheet.Cells(FirstRow, i), ActiveSheet.Cells(LastRow, i)).Select
	Selection.Borders(xlDiagonalDown).LineStyle = xlNone
	Selection.Borders(xlDiagonalUp).LineStyle = xlNone

	With Selection.Borders(xlEdgeLeft)
		.LineStyle = xlContinuous
		.ColorIndex = 0
		.Weight = xlMedium
	End With

	With Selection.Borders(xlEdgeTop)
		.LineStyle = xlContinuous
		.ColorIndex = 0
		.Weight = xlMedium
	End With

	With Selection.Borders(xlEdgeBottom)
		.LineStyle = xlContinuous
		.ColorIndex = 0
		.Weight = xlMedium
	End With

	With Selection.Borders(xlEdgeRight)
		.LineStyle = xlContinuous
		.ColorIndex = 0
		.Weight = xlMedium
	End With

	Selection.Borders(xlInsideVertical).LineStyle = xlNone
	Selection.Columns.AutoFit

End Sub