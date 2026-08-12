Attribute VB_Name = "Module5"
Sub CreateOptimizationSheet()
    Dim wsOpt As Worksheet, wsAin As Worksheet, wsBin As Worksheet
    Dim wbUser As Workbook
    Dim lastRowAin As Long, lastRowBin As Long, numRows As Long
    Dim i As Long, j As Long, rowCounter As Long
    Dim headerNumber As Long

    Set wbUser = ActiveWorkbook

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("optimization").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOpt = wbUser.Sheets.Add
    wsOpt.Name = "optimization"

    Set wsAin = wbUser.Sheets("Ain")
    Set wsBin = wbUser.Sheets("Bin")

    lastRowAin = wsAin.Cells(wsAin.Rows.Count, "A").End(xlUp).row
    lastRowBin = wsBin.Cells(wsBin.Rows.Count, "A").End(xlUp).row

    headerNumber = (lastRowAin - 1) / 2
    numRows = lastRowAin - 1

    With wsOpt
        .Range("A2").Value = "Solution/variables"
        .Range("B1").Value = "N" & headerNumber
        .Range("C1").Value = "V" & headerNumber
        .Range("D1").Value = "M" & headerNumber
        .Range("E1").Value = "lambda"
        .Range("F1").Value = "Objective"

        .Range("B2:E2").Value = 0

        .Range("A3").Value = "ObjCoeff"
        .Range("B3").Value = 0
        .Range("C3").Value = 0
        .Range("D3").Value = 0
        .Range("E3").Value = 1
        .Range("F3").Formula = "=SUMPRODUCT($B$2:$E$2,B3:E3)"

        .Range("G4").Value = "LHS"
        .Range("H4").Value = ""
        .Range("I4").Value = "RHS"

        ' Row indices column (unchanged)
        rowCounter = 5
        For i = 1 To numRows / 2
            .Cells(rowCounter, 1).Value = i
            .Cells(rowCounter + 1, 1).Value = i
            rowCounter = rowCounter + 2
        Next i

        ' ========================================================
        ' CHANGE 1: read Ain/Bin into arrays, row-normalize, write
        ' ========================================================
        ' Why: on fine discretizations (small Deltas), the coefficients
        ' of different constraint rows span many orders of magnitude
        ' (some ~ t/Deltas ~ big, some ~ 1). Simplex LP cannot resolve
        ' the true optimum when rows differ so drastically in scale,
        ' so it returns the trivial lambda = 0 solution. Dividing each
        ' row (and its RHS) by the row's max |coefficient| makes every
        ' constraint O(1) WITHOUT changing the feasible region -
        ' the LP solution is mathematically identical.
        Dim nRowsC As Long: nRowsC = lastRowAin - 1
        Dim ainVals() As Double, binVals() As Double
        ReDim ainVals(1 To nRowsC, 1 To 4)
        ReDim binVals(1 To nRowsC)

        For i = 1 To nRowsC
            For j = 1 To 4
                ainVals(i, j) = CDbl(wsAin.Cells(i + 1, j).Value)
            Next j
            binVals(i) = CDbl(wsBin.Cells(i + 1, 1).Value)
        Next i

        Dim rowMax As Double
        For i = 1 To nRowsC
            rowMax = 0
            For j = 1 To 4
                If Abs(ainVals(i, j)) > rowMax Then rowMax = Abs(ainVals(i, j))
            Next j
            If Abs(binVals(i)) > rowMax Then rowMax = Abs(binVals(i))
            If rowMax > 0.000000000001 Then
                For j = 1 To 4
                    ainVals(i, j) = ainVals(i, j) / rowMax
                Next j
                binVals(i) = binVals(i) / rowMax
            End If
        Next i

        rowCounter = 5
        For i = 1 To nRowsC
            For j = 1 To 4
                .Cells(rowCounter, j + 1).Value = ainVals(i, j)
            Next j
            .Cells(rowCounter, 7).Formula = "=SUMPRODUCT($B$2:$E$2,B" & rowCounter & ":E" & rowCounter & ")"
            .Cells(rowCounter, 8).Value = "<="
            .Cells(rowCounter, 9).Value = binVals(i)
            rowCounter = rowCounter + 1
        Next i
        ' ========================================================
        ' End of CHANGE 1
        ' ========================================================

        .Columns("A:I").AutoFit

        With .Range("A1:I" & rowCounter - 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        .Range("B2:E2").Interior.Color = RGB(146, 208, 80)
        .Range("B5:E" & rowCounter - 1).Interior.Color = RGB(0, 112, 192)
        .Range("G5:G" & rowCounter - 1).Interior.Color = RGB(0, 112, 192)
        .Range("I5:I" & rowCounter - 1).Interior.Color = RGB(0, 112, 192)
    End With
End Sub
